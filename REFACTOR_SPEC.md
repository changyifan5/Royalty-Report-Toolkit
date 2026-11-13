# REFACTOR_SPEC.md — Mac 端优化工程 · 基础设施构建规格书

> 本子任务只交付 4 个基础设施文件，**不改动任何编号脚本（01–14）**。本文件既是事实清单，也是后续并行子任务改造编号脚本时必须逐字遵循的施工说明。
>
> 验收红线：行为等价（列号/行号/颜色/页签名/文件夹约定全不变）、每个改动脚本 `~/.venv/bin/python -m py_compile` 通过、Windows COM 路径不破坏、不碰 git、不改 `.gitignore`。

---

## 0. 本次已交付的基础设施文件

| 文件 | 状态 | 作用 |
|---|---|---|
| `excel_common.py` | 新建 | 业务常量 + 收编公共函数 + 日志/弹窗/字体/win32 助手 |
| `excel_font_color.py` | 改写 | 颜色提取层全库统一，新增 07 的 `color_hint/rgb_to_hex` 与 04 的严格阈值 |
| `report_chunk.py` | 改写 | 备份命名改时间戳；模块内 2 处 Windows 字体换 fallback |
| `REFACTOR_SPEC.md` | 新建 | 本文件 |

Python 解释器：`~/.venv/bin/python`（3.13.14，openpyxl 3.1.5）。**禁止用裸 `python3`**。

---

## 1. 重复函数清单（任务一·事实）

### 1.1 `normalize_key(value)`
**各份函数体逐字相同**：
```python
def normalize_key(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None
```
- 定义位置：`row_order_common.py:54`、`01 excel_vlookup.py:86`、`04 退回修改标记.py:87`、`14 非系统发行版权vlookup.py:63`。
- 调用点：
  - 01：`:105`（在死函数内）、`:131`（`row[key_idx - min_col]`）、`:203`（`ws.cell(...).value`）。
  - 04：`:196`（A 列键）。
  - 14：`:188`、`:192`（比对表 A/B 键）、`:297`、`:323`（主文件 A/B 列键）。
- **收编结论**：01/04/14 删本地定义，改 `from excel_common import normalize_key`。行为完全一致。

### 1.2 `find_header_last_col(sheet, header_row, limit=HEADER_SCAN_LIMIT)`
函数体一致（扫 header_row 行 1..limit，最后一个非空列）。**签名差异**：
- `row_order_common.py:61` / 01`:93` / 02`:96` / 04`:94` / 14`:70`：`(sheet, header_row, limit=HEADER_SCAN_LIMIT)`，header_row 必填。
- 07`:251`：`find_header_last_col(sheet)`，无参数，内部用全局 `HEADER_ROW=6`。
- 调用点：
  - 01`:188` `find_header_last_col(ws, HEADER_ROW_A)`（位置传参）。
  - 02`:122`、`:165` `find_header_last_col(ws, HEADER_ROW)`。
  - 04`:232` `find_header_last_col(sheet, HEADER_ROW)`。
  - 07`:276`、`:321` `find_header_last_col(ws)`。
  - 14`:278` `find_header_last_col(ws, HEADER_ROW_A)`。
- **收编结论**：公共版 `excel_common.find_header_last_col(sheet, header_row=HEADER_ROW, limit=HEADER_SCAN_LIMIT)`。07 调用 `find_header_last_col(ws)` 命中默认 header_row=6，行为一致；其余位置传参不变。

### 1.3 `read_header_map(sheet, header_row)`
函数体一致（扫 header_row 行 1..500，表头去空格→列号）。**签名差异**同 1.2（07 无参用全局 6）。
- 定义：`row_order_common.py:70`、02`:105`、04`:103`、07`:260`。
- 调用：02`:204` `read_header_map(ws, HEADER_ROW)`；04`:180`、`:302` `read_header_map(ws, HEADER_ROW)`；07`:328` `read_header_map(ws)`。
- **收编结论**：`excel_common.read_header_map(sheet, header_row=HEADER_ROW=6)`。

### 1.4 `backup_file(path)`（**命名规则已变**）
旧版全部是 `path.with_suffix(path.suffix + ".bak")` 后 `shutil.copy2`：
- 定义：01`:80`、02`:80`、04`:81`、07`:359`、09`:187`、11`:74`、12`:148`、14`:57`、`05 恢复行序.py:45`、`10 列映射填表.py:245`。
- 其中 **12:148 与 10:245 在 copy2 前先 `backup_path.unlink()`**（删旧备份），其余直接覆盖。
- **新规则**（g 项，已定案）：`report.xlsx → report.bak-YYYYMMDD.xlsx`。已在 `excel_common.backup_file(path)` 实现（同日重跑覆盖同名备份，不先 unlink，与 01/02/04/07/09/14 旧行为一致——它们本来就是 copy2 覆盖）。
- **收编结论**：01/02/04/07/09/14 删本地 `backup_file`，直接把 `excel_common.backup_file` 当作 `prepare_chunk_work(..., backup_fn=...)` 传入（函数签名 `(Path)->Path` 完全兼容）。11/05恢复 同样改用。**10/12 不在本次清单，保留本地旧 `.xlsx.bak` 命名不动**（见 §6 风险点 R1）。

### 1.5 `safe_print(text)`
各份函数体仅换行/缩进微差，逻辑完全一致：
```python
def safe_print(text):
    try: print(text)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(text.encode(enc, errors="replace").decode(enc, errors="replace"))
```
- 定义：07`:58`、09`:134`、03`:47`、`05 建档行序.py:45`、11`:56`、12`:51`、13`:71`、06`:57`、08`:67`、10`:145`。
- **收编结论**：07/09 删本地，改 `from excel_common import safe_print`。其余脚本（03/05/06/08/10/11/12/13）不在本次必改清单，可后续统一；不强制。

### 1.6 `resolve_excel_path(path, label)`
函数体完全一致（文件校验后缀；文件夹内唯一 Excel；否则 sys.exit）。
- 定义：`row_order_common.py:25`、01`:51`。
- 调用：仅 01`:156` `resolve_excel_path(FILE_B, "比对文件 B")`。
- **收编结论**：01 删本地，改 `from excel_common import resolve_excel_path`。

### 1.7 `find_last_data_row` —— **01 里的死代码，直接删**
- 01`:102` 定义：
  ```python
  def find_last_data_row(sheet, key_col, data_start_row):   # 注意参数顺序 (sheet, key_col, data_start_row)
      last = data_start_row - 1
      for row in range(data_start_row, sheet.max_row + 1):
          if normalize_key(sheet.cell(row=row, column=key_col).value) is not None:
              last = row
      return max(last, data_start_row)
  ```
- **调用点：全文件零调用**（已 grep 确认；函数体内 `:105` 的 normalize_key 是自调用，不构成外部调用）。这是死代码。
- **与 row_order_common 版的差异（仅记录，因为要删）**：① 参数顺序不同（`key_col` 在 `data_start_row` 前）；② 空数据时返回 `data_start_row`（而 row_order_common 版返回 `data_start_row - 1`）；③ 不支持 read_only 模式。
- **处理**：01 删除 `:102–:107` 整个函数。`find_last_data_row` 继续只住在 `row_order_common.py:79`（已支持 read_only / 非 read_only 双模式），**不要搬到 excel_common**。

---

## 2. 常量清单（任务一·事实）

### 2.1 `HEADER_ROW` / `DATA_START_ROW` 核实结论
**绝大多数脚本确实是 HEADER_ROW=6 / DATA_START_ROW=7**：02`:47-48`、04`:49-50`、07`:49-50`、09`:61-62`、11`:41-42`、`row_order_common.py:16-17`、13`:31`（import）、10`:61`（import HEADER_ROW）。

**必须保留本地常量、不可强行套公共值的脚本**：

| 脚本 | 常量 | 值 | 说明 |
|---|---|---|---|
| 01 | `HEADER_ROW_A=6` / `DATA_START_ROW_A=7` | 主文件 A，可套公共值 | 但 `HEADER_ROW_B=1` / `DATA_START_ROW_B=2` 是比对文件 B 的布局，**保留** |
| 14 | `HEADER_ROW_A=6` / `DATA_START_ROW_A=7` | 主文件 A | `HEADER_ROW_B=2` / `DATA_START_ROW_B=3` 是比对文件 B（表头第 2 行、数据第 3 行起），**保留** |
| 12 | `HEADER_ROWS=5` | 比较第 1–5 行表头区 | 语义不同（"前 5 行"），不是"第 6 行"，**保留** |
| 05 建档/恢复 | 从 row_order_common import 6/7 | 正确 | 已集中，不动 |

**施工约定**：`excel_common.HEADER_ROW=6 / DATA_START_ROW=7` 仅作为 01-A/02/04/07/09/11/14-A 的默认值；凡涉及"比对文件 B"或"前 N 行表头区"的本地常量一律保留。

### 2.2 `EXCEL_SUFFIXES`
全部是 `{".xlsx", ".xlsm", ".xltx", ".xltm"}`，出现于 01`:48`、02`:70`、04`:68`、07`:55`、09`:68`、11`:52`、06`:45`、13（从 file_picker_common import）。**收编**：01/02/04/07/09/11 可改 `from excel_common import EXCEL_SUFFIXES`（值相同）。

### 2.3 A～D 颜色约定
- "A～D" = 前 4 列，`COLOR_CHECK_COLS = 4`（04`:52`、07`:53`、09`:64`、11`:44`）。
- 业务含义（09 页签名）：纯红→需退回、蓝→需修改、紫→线下、橙→街版、灰底+加粗→投母。
- 写回用色值：04 `COLOR_RETURN="FF0000"`（需退回）、`COLOR_RATIO="0000FF"`（需修改词曲比例）。已收编为 `excel_common.COLOR_RETURN / COLOR_RATIO`。

---

## 3. 颜色逻辑三处并存（任务一·事实 + 统一设计）

### 3.1 现状
- **(a) `excel_font_color.py`（旧）**：已有 `font_rgb_triple`（rgb/indexed/theme+tint 全支持）、`DEFAULT_THEME_HEX`、`apply_tint_component`、`hex_to_rgb`、`openpyxl_color_rgb`、`fill_rgb_triple`、宽松阈值 `is_orange/red/purple/blue_font`、`is_gray_fill`、`is_gray_bold_cell`、`is_font_bold`。
- **(b) 07**：`DEFAULT_THEME_HEX`:75、`apply_tint_component`:91、`hex_to_rgb`:99、`font_rgb_triple`:107（与 (a) 行为等价），另有 07 独有的 `rgb_to_hex`:136、`color_hint`:140。
- **(c) 04**：`font_rgb_triple`:124（**只支持 rgb/indexed，不支持 theme**）、`is_red_font`:145（严格阈值 `r>=180,g<=100,b<=100,r>g,r>b`）、`is_blue_font`:153（严格阈值 `b>=180,r<=100,b>r,b>g`）。

### 3.2 已完成的统一（改写后的 `excel_font_color.py`）
- 提取层只留一份（(a) 的超集实现），07 的 `DEFAULT_THEME_HEX/apply_tint_component/hex_to_rgb/font_rgb_triple` 全部删除本地定义，改 import。
- 07 的 `rgb_to_hex`、`color_hint` **已移入**模块。
- **阈值不合并**：04 的严格阈值与宽松组语义不同，已另立 `is_strict_red_font` / `is_strict_blue_font`。
- **现有 import 者完全兼容**：
  - 09`:37-43` import `is_blue_font, is_gray_bold_cell, is_orange_font, is_purple_font, is_red_font` —— 名字全保留。
  - 11`:30` import `is_orange_font` —— 保留。

### 3.3 04 / 07 各自要替换成什么（后续子任务执行）
- **07**：删 `:74-150`（DEFAULT_THEME_HEX/apply_tint_component/hex_to_rgb/font_rgb_triple/rgb_to_hex/color_hint 六段），改为：
  ```python
  from excel_font_color import color_hint, font_rgb_triple, rgb_to_hex
  ```
  调用点 `:166`、`:192`、`:398`、`:405`、`:466` 不变。
- **04**：删 `:124-158`（本地 font_rgb_triple/is_red_font/is_blue_font），改为：
  ```python
  from excel_font_color import is_strict_blue_font, is_strict_red_font
  ```
  `:161-174` `detect_row_mark_type` 内 `is_red_font(font)`→`is_strict_red_font(font)`、`is_blue_font(font)`→`is_strict_blue_font(font)`。
  **已知行为差异（R2，可接受）**：模块版 `font_rgb_triple` 支持 theme 色，04 旧版对 theme 返回 None。04 参考文件 B 的红/蓝字体是显式 `FF0000/0000FF`，不依赖 theme，实测无影响；若样例里出现 theme 着色需复核。

---

## 4. report_chunk.py 现状与已做的备份改造

### 4.1 `prepare_chunk_work` 接口（保持不变）
签名（已核验）：
```python
prepare_chunk_work(source: Path, script_key: str, chunk: RowChunk,
                   backup_fn: Callable[[Path], Path]) -> tuple[Path, Path]
```
返回 `(读写路径, 备份路径)`。**后续子任务不要改这个签名**，把 `excel_common.backup_file` 直接作为 `backup_fn` 传入即可。

### 4.2 已完成的备份命名改造
- 新增 `timestamp_backup_path(source)`（report_chunk 内）：`report.xlsx → report.bak-YYYYMMDD.xlsx`。
- 原 `:115`（单段非首段，实际不可达）与原 `:148`（多段非首段）两处 `source.with_suffix(source.suffix + ".bak")` 已改为 `timestamp_backup_path(source)`。
- **实际真正复制备份仍只发生在 `backup_fn(source)` 被调用时**（多段首段 `:120`、单段首段 `:113`）；后续段返回的 backup_path 仅用于打印提示。分段机制（`.chunkwork` / `.{script}.chunkprogress.json`）一行未动。

### 4.3 其他关键事实
- `fast_sheet_last_row(path, sheet_index=0)` `:238`：zip 直读 `xl/workbook.xml` + rels 定位 sheet xml，读 `<dimension ref>`。**09 的 e 项（`confirm_replace_sheet` 别再 load_workbook 看页签名）应参考它**，或直接用新的 `excel_common.list_sheet_names_fast(path)`。
- `count_data_rows(path, data_start_row=7, sheet_index=0)`、`load_file_chunks`、`compute_chunks`、`RowChunk`、`row_in_chunk`、`prompt_row_chunk`、`prompt_next_chunk_after`、`finalize_chunk_output` 全部保持。
- 模块内 2 处 Windows 字体（原 `:475`/`:478`）已换成 `tk_ui_font(10)` / `tk_mono_font(10)`（Mac 自动 PingFang/Menlo）。

---

## 5. 脚本入口与 I/O 清单（任务一·事实，供样例验证）

| 脚本 | 入口 | 选择方式 | 输入/输出约定 |
|---|---|---|---|
| 01 | `run_vlookup()`（无 argv） | tk 选主文件 A；B 固定读 `01-vlookup-比对表/`（目录内需唯一 xlsx） | 写回 A 第一个页签；无 CLI |
| 02 | `run_fix()` | tk 选 A | 写回 A；无 CLI |
| 03 | `run_split()` | tk 选 1 个 xlsx | 写回原文件；无备份（见 R3） |
| 04 | `run_mark()` | tk 选 A + B | 写回 A；无 CLI |
| 05 建档 | `run_init()` | tk 选文件 | 写回原文件 + 复制到 `00-原始基准/` |
| 05 恢复 | `run_restore()` | tk 选文件 | 写回原文件（先备份） |
| 07 | `run_split()` | tk 选文件 + 颜色/比例/表头 3 次 askstring | 写回原文件 |
| 08 | `main()`（argparse） | `--dir/--months/--preview/-y` | 不碰业务 xlsx |
| 09 | `main()` | tk 选文件 + 选 5 步骤之一 | 写回原文件，新增/重建页签 |
| 11 | `run_split()` | tk 选文件 | 写回原文件，新增「街版」页签 |
| 14 | `run_vlookup()` | tk 选 A + B + 选 B 的页签 | 写回 A |

**无头可测性**：以上除 08 外都强依赖 tkinter 文件选择器，**无法纯 CLI 无头跑主流程**。样例验证方案见 §7（用 openpyxl 直接构造样例 xlsx，再在测试脚本里绕过选择器、直接调内部函数）。

---

## 6. 交叉引用 / win32 分支 / 字体 / 日志 事实

### 6.1 import 关系（已 grep）
- `excel_font_color`：仅 09、11 import。
- `excel_sheet_copy`：03（`copy_sheet_within_workbook`）、06（`copy_sheet_between_workbooks,reorder_workbook_sheets`）、09（`copy_header_layout,copy_row_range`）、11（同 09）、12（`restore_header_region`）。
- `file_picker_common`：00/01/02/03/04/05×2/06/07/09/10/11/12/13。
- `intermediate_common`：仅 12。
- `report_chunk`：01/02/04/07/09/14。
- `row_order_common`：03/05×2/09/10/11/13。
- **编号脚本之间无互相 import**。

### 6.2 win32 分支（j 项）
| 脚本 | 分支位置 | COM 命令 | manifest / log |
|---|---|---|---|
| 03 | `:158` `if win32 and ps1.exists()` | `powershell -NoProfile -ExecutionPolicy Bypass -File 03_split_com.ps1 -Path ... -SheetRecording 录音 -SheetLyrics 词曲 -DataStartRow 7` | log `03_com_error.log`；**无 manifest**；编码 utf-8 固定。注意：`03_split_com.ps1` 当前仓库里**不存在**，实际走 Mac 分支 |
| 05 建档 | `:163` | `-File 05_init_seq_com.ps1 -Path -HeaderRow -DataStartRow -SeqHeader -SheetIndex` | log `05_com.log`；编码 **win32=gbk** |
| 06 | `:384` | `-File 06_export_sheets_com.ps1 -ManifestPath 06_export_manifest.json` | log `06_com.log`；编码 gbk；**06 不在本次必改清单** |
| 09 | `:476` | `-File 09_split_return_com.ps1 -ManifestPath 09_return_manifest.json` | log `09_com.log`；编码 **utf-8 固定（不是 gbk！）** |
| 11 | `:427` | `-File 11_split_street_com.ps1 -ManifestPath 11_street_manifest.json` | log `11_com.log`；编码 gbk |

**统一助手**：`excel_common.run_powershell(ps1, extra_args, log_file, encoding=None)`（默认 gbk/win32）。
**施工约束（逐字等价）**：03、09 调用时**必须显式传 `encoding="utf-8"`** 保持现状；05/11 用默认 gbk。`extra_args` 由各脚本照旧组装（manifest 文件名、参数名一字不改）。

### 6.3 tkinter Windows 硬编码字体（k 项）
| 位置 | 现状 |
|---|---|
| report_chunk`:475/:478` | **已改**（本任务） |
| 09`:154`（YaHei 10）、`:158`（Consolas 10） | 待改：`font=tk_ui_font(10)` / `tk_mono_font(10)` |
| 07`:205`（Consolas 10） | 待改 `tk_mono_font(10)`（07 无 YaHei 行） |
| 14`:115`（YaHei 10）、`:122`（Consolas 10） | 待改 |
| 10`:481/:484/:575/:644/:647/:655` | **10 不在清单，不动**（记录备查） |

助手：`excel_common.tk_ui_font(size, weight=None)`、`tk_mono_font(size)`，Mac 自动 PingFang SC / Menlo。

### 6.4 现有日志写法（08 先例）
08：`LOG_FILE=BASE_DIR/"08_cleanup.log"`；`append_log(lines)` 写 `\n=== YYYY-MM-DD HH:MM:SS ===\n` + 行，append 模式 utf-8。无 logging 模块。
已收编为 `excel_common.append_log(lines, log_file)`（逐字同 08）+ `excel_common.log_error(msg, log_file)`。弹窗助手 `excel_common.show_error_popup(title, message)`：tk showerror，无 GUI 静默降级。

### 6.5 写回原文件的脚本（g 项范围）
真覆盖原 xlsx 的：01、02、04、07、09、14（走 report_chunk）、03、05建档、05恢复、11、12、10。
**本次必改备份命名的**：01/02/04/07/09/14（走 report_chunk，换 `excel_common.backup_file`）+ 11、05恢复（本地 backup_file）。
**保留旧 `.xlsx.bak` 不动的**：03（本就无备份）、05建档（无备份）、10、12、06。

---

## 7. 每个编号脚本的具体改造点

> 原则：只删重复定义、换 import、加进度打印/日志/弹窗；**列号行号页签名文件夹约定全不变**。

### 01 excel_vlookup.py
1. 删 `:51-77 resolve_excel_path`、`:80-83 backup_file`、`:86-90 normalize_key`、`:93-99 find_header_last_col`、`:102-107 find_last_data_row`（**死代码，整段删**）。
2. 顶部加：`from excel_common import backup_file, find_header_last_col, normalize_key, resolve_excel_path`。
3. c 项 VLOOKUP 提速：`:202-211` 现在逐行 `ws.cell` 读 A 键 + 逐格写。改为：先用 `ws.iter_rows(min_row=row_start, max_row=row_end, min_col=key_col_a, max_col=key_col_a, values_only=True)` 批量读键，再对命中行批量写结果列（仍逐行 dict 查 lookup，但去掉每个结果格的 `ws.cell(...)` 调用，改为缓存 start_col..end_col 的 cell 对象后赋值）。**matched 计数、#N/A 填充、header 写入全部不变**。
4. h 项：长任务在 `:200` 后每 1 万行打印一次进度；关键 sys.exit 前调 `show_error_popup` + `append_log`。
5. `prepare_chunk_work(path_a, "01", chunk, backup_file)` 不变（backup_file 现在来自 excel_common，自动时间戳命名）。

### 02 半自动除错.py
1. 删 `:80-83 backup_file`、`:96-102 find_header_last_col`、`:105-112 read_header_map`。
2. import：`from excel_common import backup_file, find_header_last_col, read_header_map`（normalize_key 02 未用，不引）。
3. d 项正则合并：`:72-77` 四条 `ERROR_PATTERNS` + `:86-93 classify_error` 每行跑 4 次 search → 合并成**一条带命名分组的正则**一次 match：
   ```python
   _ERROR_RE = re.compile(
       r"(?P<找不到演唱人>找不到名稱為.+的演唱人或藝人)"
       r"|(?P<找不到歌曲>找不到歌曲名稱為.+的歌曲)"
       r"|(?P<找不到专辑>找不到專輯名稱為.+的專輯)"
       r"|(?P<收益时间错误>收益時間.+不符合專輯上、下架（月底）區間)"
   )
   ```
   `classify_error(text)` 改为 `m=_ERROR_RE.search(text.strip()); return m.lastgroup if m else None`。**匹配语义完全等价**（原来按列表顺序返回首个命中；命名分组正则也是首个命中组）。
4. h 项：主循环 `:231` 每 1 万行打印进度。

### 03 拆分词曲.py（保守）
- 已 import `find_last_data_row` from row_order_common，无需动重复函数。
- i 项（页签复制保守优化）：`run_openpyxl_split` `:139-143` 逐行逐格写，可改为行块批量；**优先级低**。
- R3：03 docstring 说"先备份 .bak"但代码**根本没备份**（`shutil` import 了却没用）。本次不补备份（行为等价优先），仅在 spec 记录。

### 04 退回修改标记.py
1. 删 `:81-84 backup_file`、`:87-91 normalize_key`、`:94-100 find_header_last_col`、`:103-109 read_header_map`、`:124-158`（本地颜色三件套）。
2. import：`from excel_common import backup_file, find_header_last_col, normalize_key, read_header_map`；`from excel_font_color import is_strict_blue_font, is_strict_red_font`。
3. `:166-168` `is_red_font/is_blue_font` → `is_strict_red_font/is_strict_blue_font`。
4. h 项进度/日志同 02。

### 05 建档行序.py
- 已用 row_order_common，无重复函数。**f 项不属于它**（f 在 05 恢复）。win32 分支 `:163` 可后续换 `run_powershell`（encoding 默认 gbk 正确），非必须。

### 05 恢复行序.py
1. f 项：`load_baseline_seqs` `:67-70` 手写循环找末行 → 改用 `row_order_common.find_last_data_row(ws, DATA_START_ROW)`；`compare_baseline` `:148-151` 同样替换。**注意这两个函数在 read_only 模式下用**，row_order_common.find_last_data_row 已支持 read_only 分支，行为等价（空数据返回 data_start_row-1，原手写也是 `last=DATA_START_ROW-1` 初始）。
2. 删 `:45-48 backup_file`，import `from excel_common import backup_file`（时间戳命名）。

### 07 G列比例拆分.py
1. 删 `:58-62 safe_print`、`:74-150` 颜色提取六件套、`:251-257 find_header_last_col`、`:260-266 read_header_map`、`:359-362 backup_file`。
2. import：`from excel_common import backup_file, find_header_last_col, read_header_map, safe_print`；`from excel_font_color import color_hint, font_rgb_triple, rgb_to_hex`。
3. k 项：`:205` `font=("Consolas",10)` → `font=tk_mono_font(10)`。
4. h 项：`scan_chunk_colors`/主循环加进度打印（已有 safe_print）。

### 09 半自动拆分.py
1. 删 `:134-140 safe_print`、`:187-190 backup_file`。import 对应。
2. e 项：`list_sheet_names` `:424-428` / `confirm_replace_sheet` `:431` / `verify_result` `:510` 用 load_workbook(read_only) 只为看页签名 → 改 `excel_common.list_sheet_names_fast(path)`（zip 直读 workbook.xml，参考 report_chunk.fast_sheet_last_row）。
3. k 项：`:154`→`tk_ui_font(10)`、`:158`→`tk_mono_font(10)`。
4. i 项：openpyxl 分支 `:412-416` 已用 `copy_row_range` 行块复制，保持。
5. win32：`:476` 分支保留；`run_com_split` 的 subprocess 可换 `run_powershell(PS1_SPLIT, ["-ManifestPath", str(MANIFEST_FILE.resolve())], LOG_FILE, encoding="utf-8")`（**显式 utf-8**，与现状一致）。

### 11 街版拆分.py
1. 删 `:56-62 safe_print`、`:74-77 backup_file`；import excel_common。
2. 已用 row_order_common 的 find_header_last_col/read_header_map/find_last_data_row，不重复。
3. win32 `:427` 可换 run_powershell（默认 gbk 正确）。

### 14 非系统发行版权vlookup.py
1. 删 `:57-60 backup_file`、`:63-67 normalize_key`、`:70-76 find_header_last_col`；import excel_common。
2. c 项提速：`:296-305`、`:322-327` 同 01 改为批量读键 + 批量写。
3. k 项：`:115`→`tk_ui_font(10)`、`:122`→`tk_mono_font(10)`。
4. `list_sheet_names` `:86-90` 可换 `list_sheet_names_fast`（优化项，非必须）。

### 不改动的脚本
- **00 / 06 / 08 / 10 / 12 / 13**：不在清单。复核其 import 未被本任务破坏：
  - 08 有用户未提交在途改动，**绝对不碰**。
  - 12 `:65` 读旧 `.xlsx.bak`（R1）——因 12 不改备份命名，该查找仍有效，不要把 12 的 backup_file 换成时间戳版。
  - 10/12/13 import 的 row_order_common / file_picker_common / excel_sheet_copy 名字均未变，无破坏。

---

## 8. 行为等价样例验证方案（后续子任务照此跑）

统一在仓库根目录用 `~/.venv/bin/python`，构造小 xlsx 后**直接调内部函数**绕过 tkinter 选择器：

1. **01**：造 `主.xlsx`（前 6 行表头、第 6 行有 A 列表头、第 7 行起 A 列键）+ `01-vlookup-比对表/比.xlsx`（B 列键、B/D/E/F 列值）。直接调 `load_lookup_table` + 手工拼 RowChunk(1,1,7,N,…) 跑 `run_vlookup` 核心循环。期望：主表末 4 列 = B/D/E/F 命中值，未命中 `#N/A`；页签名不变；生成 `主.bak-YYYYMMDD.xlsx`。
2. **02**：在 01 产物上跑 `classify_error` 单测（喂 4 条典型比對結果串，断言返回名），再跑主循环断言 B/C/D 被覆盖 + 标黄 + 已完成列。
3. **04**：造 B 文件 A～D 列分别纯红/纯蓝字体行，断言 `detect_row_mark_type` 返回 需退回/需修改词曲比例（用 strict 函数，阈值必须与旧版逐值一致）。
4. **07**：造 G 列数值 + A～D 某 rgb 字体行，跑 `split_value(total, r1, r2)` 断言 v1+v2==total。
5. **09**：造含红/蓝/紫/橙/灰底加粗行的表，跑 `cell_matches_mode` 断言各 scan_mode 命中集；用 `list_sheet_names_fast` 断言页签名与 openpyxl 一致。
6. **14**：同 01，额外断言 B 列二次匹配只回填 B~D 三列。
7. **通用**：每个改过的脚本先 `~/.venv/bin/python -m py_compile "NN xxx.py"`。

## 9. 可直接单测的纯函数（为 m 项铺路）
- excel_common：`normalize_key`、`find_header_last_col`、`read_header_map`、`timestamp_backup_path`、`safe_print`、`tk_ui_font/tk_mono_font`。
- excel_font_color：`hex_to_rgb`、`apply_tint_component`、`rgb_to_hex`、`color_hint`、`is_strict_red/blue_font`、`is_orange/red/purple/blue_font`。
- report_chunk：`compute_chunks`、`format_chunk_label`、`row_in_chunk`、`timestamp_backup_path`、`fast_sheet_last_row`（用小 xlsx）。
- 脚本内纯函数：02 `classify_error`、`cell_text`；07 `split_value`、`parse_ratio`、`to_number`；11 `split_formula`、`calc_street_g_value`；13 `parse_i_date`、`excel_serial_to_date`。

---

## 10. 风险点 / 与预想不符的发现（R 编号）
- **R1**：12 `:65` 与 12/10 的 backup 仍是旧 `.xlsx.bak`。若后续有人误把 12/10 也换成时间戳备份，12 `:65` 的 `target.with_suffix(".xlsx.bak")` 查找会失效。**结论：10/12 保持旧命名，不要动。**
- **R2**：04 颜色提取换模块版后，theme 着色字体从"返回 None"变成"可解析"。04 实际用色为显式 RGB，无影响；样例验证时若 B 文件含 theme 色需复核。
- **R3**：03 docstring 声称"先备份 .bak"，代码实际无备份。本次不补（行为等价），仅记录。
- **R4**：01 的 `find_last_data_row` 与 row_order_common 版在**空数据返回值**和**参数顺序**上都不同——因 01 版是死代码，直接删，不做语义迁移。
- **R5**：`03_split_com.ps1` 在仓库中不存在；03 的 win32 分支 `and PS1_SCRIPT.exists()` 永远走 Mac 分支。win32 助手改造时不要"顺手创建"该 ps1。
