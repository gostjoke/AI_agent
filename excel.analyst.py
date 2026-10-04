"""excel_analyst.py — 讀取 Excel、用 pandas 計算、由模型寫總結

用法：
    python excel_analyst.py sales.xlsx
    python excel_analyst.py sales.xlsx "哪個產品營收最高？各月趨勢如何？"

設計重點：
    - 模型不直接看整份資料，也不自己算數字。所有數字都由 pandas 算好再交給模型。
    - 檔案路徑由程式綁定，模型不需要（也不能）自己傳路徑。
    - 三個工具全部唯讀，不會改動你的檔案。
"""
from __future__ import annotations

import logging
import sys
import warnings
from pathlib import Path

import pandas as pd

from agent import Agent, ToolBox

AGGS = {"sum", "mean", "count", "min", "max", "median"}
PERIODS = {"day": "D", "week": "W", "month": "M", "quarter": "Q", "year": "Y"}

SYSTEM = """你是資料分析助理，負責分析一份 Excel 檔並寫出總結。

工作流程：
1. 先呼叫 overview 了解有哪些工作表、欄位和資料型別。
2. 用 column_stats 看重要欄位的分布，用 aggregate 做分組彙總（例如各產品營收、各月趨勢）。
3. 資料足夠後，用繁體中文寫總結。

規則：
- 總結裡的每個數字都必須來自工具的回傳結果，不可以自己心算或估計。
- 欄位名稱必須和 overview 回傳的完全一致。
- 不要重複呼叫已經得到結果的查詢。

總結格式：
【資料概況】幾個工作表、各有多少筆、涵蓋的範圍
【重點發現】3 到 5 點，每點附上具體數字
【需要注意】空值、異常值、可疑的資料（沒有就寫「未發現」）
"""


def build_toolbox(path: str) -> ToolBox:
    """讀入一份 Excel，回傳綁定在這份檔案上的工具組。"""
    sheets = pd.read_excel(path, sheet_name=None)  # {工作表名稱: DataFrame}，只讀一次
    for df in sheets.values():
        df.columns = [str(c).strip() for c in df.columns]

    def _sheet(name: str) -> pd.DataFrame:
        if name not in sheets:
            raise ValueError(f"No sheet '{name}'. Available sheets: {list(sheets)}")
        return sheets[name]

    def _col(df: pd.DataFrame, col: str) -> pd.Series:
        if col not in df.columns:
            raise ValueError(f"No column '{col}'. Available columns: {list(df.columns)}")
        return df[col]

    box = ToolBox()

    @box.tool
    def overview() -> str:
        """Show every sheet in the Excel file: row count, column names, data types, empty-cell counts, and the first 5 rows. Always call this first."""
        out = []
        for name, df in sheets.items():
            out.append(f"## Sheet: {name} — {len(df)} rows x {len(df.columns)} columns")
            for c in df.columns:
                out.append(f"- {c} ({df[c].dtype}, {int(df[c].isna().sum())} empty)")
            out.append("First 5 rows:")
            out.append(df.head(5).to_csv(index=False).strip())
            out.append("")
        return "\n".join(out)

    @box.tool
    def column_stats(sheet: str, column: str) -> str:
        """Statistics for one column. Numeric: sum, mean, median, min, max. Text: the 10 most frequent values. Date: earliest and latest.

        Args:
            sheet: Sheet name, exactly as shown by overview.
            column: Column name, exactly as shown by overview.
        """
        s = _col(_sheet(sheet), column)
        out = [f"column={column} rows={len(s)} empty={int(s.isna().sum())} unique={s.nunique()}"]
        if pd.api.types.is_numeric_dtype(s):
            for k in ("sum", "mean", "median", "min", "max", "std"):
                out.append(f"{k}={round(float(getattr(s, k)()), 2)}")
        elif pd.api.types.is_datetime64_any_dtype(s):
            out.append(f"earliest={s.min()} latest={s.max()}")
        else:
            out.append("most frequent values (value,count):")
            out.append(s.value_counts().head(10).to_csv(header=False).strip())
        return "\n".join(out)

    @box.tool
    def aggregate(sheet: str, group_by: str, value_column: str, agg: str = "sum",
                  period: str = "", top_n: int = 20) -> str:
        """Group rows by one column and aggregate another, e.g. total revenue per product, or revenue per month.

        Args:
            sheet: Sheet name, exactly as shown by overview.
            group_by: Column to group by, e.g. "Product" or a date column.
            value_column: Column to aggregate, e.g. "Revenue".
            agg: One of sum, mean, count, min, max, median.
            period: Only when group_by is a date column: day, week, month, quarter, or year. Otherwise leave empty.
            top_n: Maximum number of groups to return.
        """
        df = _sheet(sheet)
        key, val = _col(df, group_by), _col(df, value_column)
        agg = str(agg).lower().strip()
        if agg not in AGGS:
            raise ValueError(f"agg must be one of {sorted(AGGS)}")
        if agg != "count" and not pd.api.types.is_numeric_dtype(val):
            raise ValueError(f"'{value_column}' is not numeric ({val.dtype}); only agg='count' works on it.")

        period = str(period or "").lower().strip()
        if period in {"none", "null", "no", "n/a"}:
            period = ""
        if period:
            if period not in PERIODS:
                raise ValueError(f"period must be one of {sorted(PERIODS)} or empty")
            # 模型有時會對非日期欄位亂填 period，這裡擋掉而不是算出沒意義的結果
            if pd.api.types.is_numeric_dtype(key):
                raise ValueError(f"'{group_by}' is not a date column; leave period empty.")
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                dates = pd.to_datetime(key, errors="coerce")
            if dates.notna().sum() < key.notna().sum() * 0.5:
                raise ValueError(f"'{group_by}' is not a date column; leave period empty.")
            key = dates.dt.to_period(PERIODS[period]).astype(str)

        result = val.groupby(key).agg(agg)
        total = len(result)
        result = result.sort_index() if period else result.sort_values(ascending=False)
        result = result.head(int(top_n)).round(2)
        header = f"{group_by},{agg}_{value_column}"
        note = f"\n(showing {len(result)} of {total} groups)" if total > len(result) else ""
        return header + "\n" + result.to_csv(header=False).strip() + note

    return box


def analyze(path: str, question: str = "分析這份資料並做總結。", **agent_kwargs) -> str:
    """給其他程式調用的入口：傳入檔案路徑和問題，回傳總結文字。"""
    agent = Agent("excel-analyst", SYSTEM, build_toolbox(path), max_steps=15, **agent_kwargs)
    return agent.run(question)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    if len(sys.argv) < 2:
        sys.exit('用法: python excel_analyst.py <檔案.xlsx> ["想問的問題"]')
    file = sys.argv[1]
    if not Path(file).exists():
        sys.exit(f"找不到檔案: {file}")
    q = " ".join(sys.argv[2:]) or "分析這份資料並做總結。"
    print("\n" + analyze(file, q))