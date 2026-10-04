# Local Agent

用 Python 驅動、跑在自己顯示卡上的 AI agent。模型由 Ollama 提供，所有能力都是你自己寫的 Python 函式。

## 專案結構

```
AI_agent/
├── agent.py           核心（ToolBox、Agent）＋ 檔案助理範例
├── excel_analyst.py   Excel 分析 agent，借用 agent.py 的核心
└── README.md
```

`agent.py` 是所有東西的基礎，要一直留著。每件新的自動化工作各自開一個檔案，像 `excel_analyst.py` 那樣 `from agent import Agent, ToolBox`。

| 元件 | 位置 | 用途 |
| --- | --- | --- |
| `ToolBox` | `agent.py` | 工具註冊表，把 Python 函式變成模型可以呼叫的工具 |
| `Agent` | `agent.py` | 一個名字 + 一段 system prompt + 一組工具，用 `run()` 執行任務 |
| 檔案助理 | `agent.py` | 現成可用：列檔案、讀寫文字檔、建立 Excel |
| Excel 分析 | `excel_analyst.py` | 現成可用：讀取 Excel、計算、寫總結 |

## 需求

- Python 3.10 以上
- [Ollama](https://ollama.com) 已安裝並在背景執行
- 16GB VRAM 的顯示卡（預設模型 `gpt-oss:20b`）
- 模型必須支援 tool calling（見[常見問題](#常見問題)）

## 安裝

```bash
ollama pull gpt-oss:20b
pip install ollama openpyxl pandas
```

`ollama` 是核心必需；`openpyxl` 用於建立 Excel；`pandas` 用於 Excel 分析。

---

## 用法一：檔案助理（agent.py）

```bash
python agent.py "列出目前資料夾的檔案，告訴我哪個最新"
python agent.py "建立 tables.xlsx，三個工作表：Sales、Inventory、Employees，每個放幾筆範例資料"
```

執行時第一行會印出模型和工具清單，接著是每一步呼叫了哪個工具，最後是模型的回答：

```
[file-helper] model=gpt-oss:20b tools=['list_files', 'read_text_file', 'write_text_file', 'write_excel_sheet']
[file-helper] step 1 → list_files({'folder': '.'})
最新的檔案是 agent.py ...
```

第一次執行會比較慢，因為模型要載入 VRAM。用 `ollama ps` 可以確認模型是否完整跑在 GPU 上。

### 內建工具

| 工具 | 做什麼 | 需要確認 |
| --- | --- | --- |
| `list_files` | 列出資料夾內的檔案、大小、修改時間 | 否 |
| `read_text_file` | 讀取 UTF-8 文字檔 | 否 |
| `write_text_file` | 寫入純文字檔（.txt .md .csv .json 等）。遇到 `.xlsx` 會拒絕 | 是 |
| `write_excel_sheet` | 在 `.xlsx` 中建立或取代一個工作表，一次呼叫寫一張表 | 是 |

「需要確認」的工具執行前會在終端機問 `y/N`，按 `y` 才會真的寫入。建立三個工作表就會問三次。

---

## 用法二：Excel 分析（excel_analyst.py）

```bash
python excel_analyst.py sales.xlsx
python excel_analyst.py sales.xlsx "哪個產品營收最高？各月趨勢如何？"
```

不給問題時會做整體分析，並依照「資料概況、重點發現、需要注意」三段寫出總結。

### 運作方式

模型不直接看整份資料，也不自己算數字。它只決定要看什麼，計算全部交給 pandas：

| 工具 | 做什麼 |
| --- | --- |
| `overview` | 列出所有工作表、欄位、型別、空值數、前 5 筆 |
| `column_stats` | 單一欄位統計：數值看總和與平均，文字看最常出現的值，日期看起訖 |
| `aggregate` | 分組彙總，例如各產品營收。日期欄位可按 day / week / month / quarter / year 彙總 |

三個工具全部唯讀，不會改動檔案，所以不需要確認，可以直接放進排程。檔案路徑由程式綁定，模型無法讀取其他檔案。

### 在其他程式裡調用

```python
from excel_analyst import analyze

summary = analyze("sales.xlsx", "各區域的銷售表現如何？")
print(summary)
```

### 限制

- **假設每個工作表的第一列是欄位名稱。** 報表上方有標題列、合併儲存格，或一個工作表放了好幾個表格時，讀出來的欄位會是亂的，需要先整理或另外加處理。
- **只能分組彙總，不能加篩選條件。** 例如「只看北區的各產品營收」目前做不到。
- **前幾次請抽查數字。** 數字是 pandas 算的，但模型可能引用錯或解讀錯。對照 log 裡每個 `step` 的查詢就能確認。

---

## 寫自己的 agent

每件工作開一個新檔案，不要改 `agent.py`。

### 1. 寫工具

```python
# my_tools.py
import pandas as pd
import pyodbc
from agent import ToolBox

box = ToolBox()

@box.tool
def run_sql(query: str) -> str:
    """Run a read-only SELECT query on the ERP database and return up to 50 rows.

    Args:
        query: A single T-SQL SELECT statement.
    """
    conn = pyodbc.connect("DSN=erp_readonly")
    return pd.read_sql(query, conn).head(50).to_markdown()

@box.tool(confirm=True)
def save_report(filename: str, content: str) -> str:
    """Save a report as a text file in the reports folder.

    Args:
        filename: File name only, without folder.
        content: Full text of the report.
    """
    path = f"reports/{filename}"
    open(path, "w", encoding="utf-8").write(content)
    return f"saved to {path}"
```

寫工具時的規則：

- **docstring 就是模型看到的說明。** 第一行寫清楚這個工具做什麼、回傳什麼。寫得越具體，模型越不會用錯。
- **每個參數都要有型別標註**，並在 `Args:` 區塊說明。型別用 `str`、`int`、`float`、`bool`、`list` 這類簡單型別。參數有固定格式時附上範例。
- **回傳字串。** 回傳值會被轉成文字交給模型，超過 8000 字元會被截斷（`MAX_TOOL_OUTPUT`），所以工具本身就該限制筆數或做摘要。
- **出錯直接 raise，並在訊息裡給出正確答案。** 例外會被接住並回報給模型。訊息寫成 `No column 'x'. Available columns: [...]` 這種形式，模型下一步就能自己修正。
- **不該做的事在函式裡擋掉。** 模型會用任何看起來能完成任務的工具硬做。`write_text_file` 拒絕 `.xlsx` 就是這個原因，用程式擋比在 prompt 裡叮嚀可靠。
- **程式已經知道的事不要讓模型傳。** 像 `excel_analyst.py` 把檔案路徑綁在程式裡，工具就沒有 `path` 參數，少一個出錯的機會。
- **數字交給程式算。** 加總、平均、排序都寫在工具裡，模型只負責解讀結果。
- **會寫入、寄出、刪除的工具加 `confirm=True`。**

### 2. 建立並調用 agent

```python
# daily_report.py
import logging
from agent import Agent
from my_tools import box

logging.basicConfig(level=logging.INFO, format="%(message)s")

report_agent = Agent(
    name="report",
    system=(
        "你是庫存報表助理。先用 run_sql 查詢需要的資料，"
        "再整理成繁體中文摘要。不確定欄位名稱時先查 INFORMATION_SCHEMA。"
    ),
    toolbox=box.subset("run_sql", "save_report"),
)

print(report_agent.run("列出上週庫存低於安全庫存的料號，存成 low_stock.txt"))
```

每次 `run()` 互相獨立，不會記得上一次的對話。需要延續的資訊要寫進任務描述裡，或是存成檔案讓工具去讀。

### Agent 參數

| 參數 | 預設值 | 說明 |
| --- | --- | --- |
| `name` | 必填 | 顯示在 log 裡，用來分辨是哪個 agent |
| `system` | 必填 | 角色與做事規則 |
| `toolbox` | 必填 | 這個 agent 可以用的工具 |
| `model` | `gpt-oss:20b` | Ollama 的模型名稱，必須支援 tools |
| `num_ctx` | `32768` | context 長度。16GB VRAM 不建議再調高 |
| `max_steps` | `12` | 最多來回幾輪，防止無限迴圈 |
| `host` | `None` | Ollama 位址。模型在另一台機器時填 `http://<ip>:11434` |
| `approve` | 終端機詢問 | `confirm=True` 的工具執行前呼叫的確認函式 |

### 一個 agent 只給它需要的工具

工具越多，模型越容易選錯。用 `subset()` 從同一個 `ToolBox` 切出不同組合：

```python
reader = Agent("reader", "...", box.subset("run_sql"))
writer = Agent("writer", "...", box.subset("run_sql", "save_report"))
```

單一 agent 建議控制在 5 個工具以內。

## 確認機制

標記為 `confirm=True` 的工具，執行前會呼叫 `approve(name, args)`，回傳 `True` 才會真的執行。

```python
# 預設：在終端機問 y/N
Agent("a", "...", box)

# 無人值守（排程）：一律拒絕危險動作
Agent("a", "...", box, approve=lambda name, args: False)

# 自訂規則：只允許寫入 reports 資料夾
def only_reports(name, args):
    return name == "save_report" and "/" not in args.get("filename", "")

Agent("a", "...", box, approve=only_reports)
```

被拒絕時模型會收到 `DENIED`，並在最終回答中說明沒有完成的部分。

## 排程執行

把任務寫成一支腳本，再交給作業系統排程。

Windows 工作排程器，動作指向一個 `.bat`：

```bat
cd /d C:\path\to\project
call .venv\Scripts\activate
python excel_analyst.py data\weekly.xlsx >> logs\weekly.log 2>&1
```

Linux / macOS 的 cron：

```
0 8 * * 1-5  cd /path/to/project && python daily_report.py >> logs/daily_report.log 2>&1
```

排程執行沒有人可以按 y。唯讀的 agent（像 Excel 分析）可以直接排程；有 `confirm=True` 工具的 agent 要把 `approve` 換成自動規則，否則程式會卡在等待輸入。

## 常見問題

**`does not support tools (status code: 400)`**
選的模型不支援 tool calling（例如 `gemma3:4b`）。換成支援的模型，用 `ollama show <模型名稱>` 檢查，Capabilities 裡要有 `tools`。

**`ModuleNotFoundError: No module named 'agent'`**
`agent.py` 不在同一個資料夾。所有 agent 檔案都要和 `agent.py` 放在一起。

**連線被拒絕（Connection refused）**
Ollama 沒有在執行。開啟 Ollama 應用程式，或在終端機執行 `ollama serve`。

**跑得很慢**
執行 `ollama ps`，看 PROCESSOR 欄位。如果不是 100% GPU，代表模型加上 context 超過 VRAM，一部分掉到 CPU 了。調低 `num_ctx` 或換小一點的模型。

**產生的 Excel 打開是亂碼或純文字**
模型用錯了工具。確認第一行 log 的 `tools=` 裡有 `write_excel_sheet`，沒有的話代表 `agent.py` 是舊版。

**Excel 分析的欄位名稱變成 `Unnamed: 0` 之類**
工作表的第一列不是欄位名稱。見 Excel 分析的[限制](#限制)。

**模型沒有呼叫工具，直接亂答**
通常是 docstring 太模糊，或 system prompt 沒有要求它使用工具。在 system 裡明確寫「必須先用某某工具取得資料，不可以自行假設」。

**模型呼叫工具時參數錯誤**
檢查型別標註與 `Args:` 說明是否完整。參數有固定格式（日期、料號）時，在說明裡附上範例。

**回答「超過 N 步仍未完成」**
任務太大或模型在繞圈。把任務拆小，或是把多個步驟合併成一個功能更完整的工具。

**工具結果後面出現 `...[truncated]`**
回傳內容超過 8000 字元。讓工具自己限制筆數、只回傳需要的欄位，或先做彙總再回傳。

**結果不如預期，不知道哪裡錯**
看 log 裡的 `step N → 工具名稱(參數)`。從那幾行可以直接看出模型選了哪個工具、傳了什麼參數、在哪一步走偏。

## 安全注意事項

- 資料庫連線使用唯讀帳號，不要讓模型能執行任意寫入的 SQL。
- 不要提供可以執行任意 shell 指令或任意 Python 程式碼的工具。
- 檔案類工具限制在特定資料夾內，在函式裡檢查路徑。
- 模型的輸出可能有錯。會影響正式資料的動作一律走 `confirm=True`，報表類結果在信任之前先人工抽查幾次。

## 下一步

1. 拿實際會用到的 Excel 跑 `excel_analyst.py`，確認總結的品質。
2. 挑一件每週重複多次、只需要讀資料的工作，照 `excel_analyst.py` 的模式開一個新檔案，為它寫 2 到 3 個工具。
3. 手動跑穩之後再放進排程。
4. 需要給其他人使用時，再把 `run()` 包進 Django（背景 worker 執行，併發數設為 1）。