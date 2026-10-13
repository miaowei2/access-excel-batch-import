# Access 通用 Excel 批量导入 / Access Universal Excel Batch Import

[中文](#中文) | [English](#english)

---

<a id="中文"></a>

## 中文

### 简介

纯 VBA 实现的 Access 通用 Excel 批量导入工具。运行一个过程即可自动生成导入窗体，按“目标表 → Excel 文件 → 工作表 → 字段映射 → 开始导入”的流程，把 Excel 数据批量追加到 Access 本地表、Access 后端链接表或 SQL Server ODBC 链接表。

### 功能特点

- 支持 `.xls`、`.xlsx`、`.xlsm`、`.xlsb`；只读打开，不更新外部链接，禁用源工作簿宏。
- 可选择工作表和标题行（默认第 1 行），按字段名或 Caption 自动匹配，也可手动映射或跳过列。
- 类型校验：文本、数字、货币、Decimal、bigint、日期、是/否，以及长度和必填项检查。
- 默认整批校验：有任何错误就不写入目标表；也可勾选“跳过校验失败的行”，只追加通过校验的行。
- 错误明细记录 Excel 原始行号、字段、原始值和原因；批次历史记录源文件、工作表、行数、追加数和状态。
- 目标追加在一个数据库事务内完成，失败整批回滚。
- 导入过程可取消；界面提示支持中英文切换。
- 不依赖 Excel 驱动猜测列类型，文本编号、长备注、引号、换行等内容原样保留。

### 文件

| 文件 | 说明 |
| --- | --- |
| `modExcelImport.bas` | 核心模块：文件选择、字段映射、读取 Excel、校验、追加、取消、语言切换 |
| `modExcelImportForms.bas` | 安装配置表、创建暂存表、自动生成四个窗体 |
| `.vscode/settings.json` | VS Code 以 GBK 编码打开 `.bas` / `.cls` |

> **编码提示：** 两个 `.bas` 文件为 **CP936 / GBK 编码、无 `Attribute` 头**，源码中直接包含中文。请在中文 Windows 环境下通过 VBE 的“导入文件”加载，不要转换为 UTF-8，也不要用 `LoadFromText` 加载。GitHub 网页上预览这两个文件时中文会显示为乱码，属于正常现象，请下载后使用。

### 环境要求

- Windows 中文系统，Access 2010 及以上（32 / 64 位），已安装 Excel。
- 数据库所在目录为受信任位置，或已启用该数据库的内容。
- VBA 引用 **Microsoft Office xx.0 Access database engine Object Library**（DAO，仅用于读取表结构）。ADO、Excel 使用后期绑定，无需额外引用。

### 安装

1. 在备份的数据库中按 `Alt+F11` 打开 VBE，通过“文件 → 导入文件”导入 `modExcelImport.bas` 和 `modExcelImportForms.bas`。两个文件需同时导入；如已存在同名模块，先移除再导入。
2. “工具 → 引用”中勾选 DAO 引用（见上）。
3. 执行“调试 → 编译”，确认无编译错误。
4. 在立即窗口执行：

```vba
CreateExcelImportForm
```

将自动创建以下对象：

| 对象 | 说明 |
| --- | --- |
| `frmExcelImport` | 主导入窗体 |
| `frmExcelImportMap` | 字段映射子窗体 |
| `frmExcelImportErrors` | 当前批次错误明细（只读） |
| `frmExcelImportHistory` | 导入历史（只读） |
| `tblEI_Map` | 当前字段映射 |
| `tblEI_Batches` | 批次记录 |
| `tblEI_Errors` | 错误明细 |

重复执行会重建并覆盖四个窗体（需先关闭它们），不会删除业务表和导入历史。如需自定义窗体，请先另存备份。

只创建而不打开：

```vba
CreateExcelImportForm True
```

### 使用方式

1. 打开窗体：

```vba
DoCmd.OpenForm "frmExcelImport"                          ' 手动选择目标表
DoCmd.OpenForm "frmExcelImport", , , , , , "客户"        ' 预设并锁定目标表
```

2. 选择目标表（或通过 `OpenArgs` 预设）。
3. 选择 Excel 文件（预设目标表时打开窗体会自动弹出文件对话框），再选择工作表和标题行。
4. 在字段映射表中检查或调整“目标字段”，不需要导入的列勾选“跳过”。目标字段不允许重复映射。
5. 需要时勾选“跳过校验失败的行”，点击“开始导入”。
6. 导入结束后查看提示；如有错误，可打开错误明细和导入历史。

### 中英文切换

界面默认英文。点击页脚的“中文 / English”按钮切换，或执行：

```vba
EI_SetLanguage "zh"   ' 中文
EI_SetLanguage "en"   ' English
```

设置保存在当前数据库的 `EI_Language` 项目属性中。已写入日志表的错误原因保持写入时的语言。

### 数据转换规则

- 空单元格转换为 Null，不替换为 0；映射列全空的行不追加。
- 文本原样保留。若 Excel 中的编号本身是数字、仅用显示格式补零，则无法恢复前导零，请先将源列设为文本。
- 布尔值：`否 / FALSE / NO / 0` 为 False；`是 / TRUE / YES / 1 / -1` 为 True。
- 日期：支持 Excel 序号（1900 / 1904 日期系统）以及 `2026.10.12` 形式的文本；其他日期文本按本机地区格式解释，建议使用明确的年月日格式。
- 整数字段不接受小数，溢出记为错误；bigint 使用 Decimal 检查 64 位范围；Decimal 受 VBA Decimal 约 28 位精度限制。
- Excel 公式错误值记为错误（内部以保留标记 `#EI_EXCEL_ERROR#` 标识，请勿作为正常源文本）。
- 不导入自动编号 / identity、计算列、rowversion、附件、OLE、多值字段、GUID 等不支持的字段；未映射字段由目标库应用默认值。

### 实现方式

```
Excel 工作表
   │  Range.Value2 按块读取（不逐单元格读取）
   ▼
ADO 客户端缓存 ── UpdateBatch ──► 暂存表 tblEI_Stage（长文本列）
                                       │  ADO Recordset 分块读取，VBA 内存中校验类型 / 长度 / 必填
                                       ▼
                                  错误明细 tblEI_Errors
                                       │  通过校验的数据
        ┌──────────────────────────────┴──────────────────────────────┐
        ▼                                                             ▼
Access 目标：生成类型化暂存表                            SQL Server 目标：≤500 行 / ≤2000 参数分批
tblEI_TypedStage（UpdateBatch）                          写入连接专属临时表 #EI_Stage
        │  同一连接、同一事务                                          │  同一连接、同一事务
        ▼                                                             ▼
INSERT INTO 目标表 SELECT … FROM 暂存表                  INSERT INTO 目标表 SELECT … FROM #EI_Stage
```

- **读取：** 按块一次读取 `Range.Value2`，只映射被选中的列，不使用 CSV / 文本驱动。
- **暂存：** 所有映射列先以文本 / 长文本形式写入暂存表，避免类型被猜测；Excel 行号单独保存用于定位错误。
- **校验：** 在 VBA 内存中逐字段转换并检查，不在 SQL 中调用 VBA 自定义函数，因此也适用于外部 ADO / SQL Server。
- **追加：** Access 目标由类型化暂存表经一条 `INSERT INTO … SELECT` 追加；SQL Server 目标直连 ADO，通过 `#EI_Stage` 临时表在服务器端追加。使用 ADO `BeginTrans / CommitTrans / RollbackTrans`，提交前核对实际追加行数。
- **数据库访问：** 本地 Access 操作统一使用 `CurrentProject.Connection`；DAO 仅用于读取表 / 字段定义。批次日志使用 `ADODB.Command` 参数化写入，文本参数使用 Unicode 类型。
- **Excel 实例：** 只关闭模块自己创建的 Excel 实例。
- **窗体：** 窗体、控件及事件代码均由 `modExcelImportForms` 在运行时通过 VBA 生成。

### 失败处理与取消

- 默认整批校验：存在类型 / 必填 / 长度错误时不写入目标表。
- 勾选“跳过校验失败的行”后，仅追加通过校验的行。
- 唯一索引、主键、关系、验证规则、精度等最终约束由数据库执行；违反时目标追加事务整批回滚，错误以 `ExcelRow=0` 记录为批次级错误。
- 提交或回滚时连接异常，批次标记为“提交状态待核实”，请核实目标表后再决定是否重试。
- 导入期间“关闭”按钮变为“取消导入”；分块读取及字段校验之间可响应取消。单条 SQL 执行期间无法中断，提交完成后不可取消。

### SQL Server 连接

默认从链接表的 `TableDef.Connect` 读取 ODBC 配置，通过 `Provider=MSDASQL` 直连后端。要求 ODBC 驱动 / DSN 与 Access 位数一致，当前账号具备读取表结构、追加数据和创建 tempdb 临时表的权限。

如果链接未保存密码，或需使用 SQL Server OLE DB 驱动，可在运行期间显式指定连接：

```vba
' "客户" 是 Access 中的链接表名，不是 SQL Server 中的表名
EI_SetSqlServerConnection "客户", "Provider=MSOLEDBSQL;Data Source=YOUR_SERVER;Initial Catalog=YOUR_DB;Integrated Security=SSPI;"
```

连接字符串只保存在当前 VBA 会话内存中，不写入源码或日志；建议使用 Windows 身份验证，不要把真实密码写进代码。传入空字符串则恢复使用链接表的 ODBC 配置。更换连接后需重新选择工作表 / 目标表，或调用 `EI_Map`。SQL Server 中文文本推荐使用 `nvarchar` / `ntext`。

### 限制

- 支持本地 Access 表、Access 后端链接表和 SQL Server ODBC 链接的用户表；不支持其他 ODBC 数据库、SQL Server 视图、跨库 / 四段表名及 Excel 等异构链接目标。
- 一次最多映射 254 列。
- 多人使用时每人需要独立的前端文件，不要共用同一前端中的映射 / 暂存表。
- 大批量导入会增长 Access 文件，请预留磁盘空间，必要时压缩修复。
- 本项目尚未在所有环境（Access 版本、位数、SQL Server 驱动）中验证，使用前请在备份数据库中测试。

### 开源协议

[MIT License](LICENSE)

---

<a id="english"></a>

## English

### Overview

A pure-VBA, generic Excel batch import tool for Microsoft Access. Run one procedure to generate the import forms, then follow the familiar flow **target table → Excel file → worksheet → field mapping → import** to append Excel data into a local Access table, a linked Access back-end table, or a SQL Server table linked through ODBC.

### Features

- Supports `.xls`, `.xlsx`, `.xlsm`, `.xlsb`; workbooks are opened read-only, external links are not updated, and macros in the source workbook are disabled.
- Choose the worksheet and header row (default: row 1); columns are matched automatically by field name or Caption, and can be remapped or skipped manually.
- Validation of text, number, currency, Decimal, bigint, date and Yes/No values, plus length and required-field checks.
- Whole-batch validation by default: any error prevents writing to the target table. Optionally, "skip rows that fail validation" appends only valid rows.
- Error details store the original Excel row, field, raw value and reason; batch history stores the source file, worksheet, row counts and status.
- The append runs in a single database transaction and is rolled back as a whole on failure.
- Cancelable import; UI messages can be switched between English and Chinese.
- No column-type guessing by an Excel driver; text IDs, long memos, quotes and line breaks are preserved as-is.

### Files

| File | Description |
| --- | --- |
| `modExcelImport.bas` | Core module: file picker, field mapping, Excel reading, validation, append, cancel, language switch |
| `modExcelImportForms.bas` | Installs the config tables, creates staging tables, and generates the four forms |
| `.vscode/settings.json` | Opens `.bas` / `.cls` files as GBK in VS Code |

> **Encoding note:** both `.bas` files are **CP936 / GBK encoded without an `Attribute` header** and contain Chinese text directly in the source. Import them through the VBE "Import File" command on a Chinese (CP936) Windows system. Do not convert them to UTF-8 and do not load them with `LoadFromText`. They appear garbled in the GitHub web preview; this is expected, so download the files instead.

### Requirements

- Windows with a Chinese (CP936) system code page, Access 2010 or later (32 / 64-bit), Excel installed.
- The database must be in a trusted location, or its content must be enabled.
- VBA reference to **Microsoft Office xx.0 Access database engine Object Library** (DAO, used only to read table definitions). ADO and Excel are late-bound and need no references.

### Installation

1. In a backup copy of your database, press `Alt+F11` to open the VBE and use "File → Import File" to import `modExcelImport.bas` and `modExcelImportForms.bas`. Import both together; if modules with the same names exist, remove them first.
2. In "Tools → References", tick the DAO reference above.
3. Run "Debug → Compile" and make sure there are no compile errors.
4. In the Immediate window, run:

```vba
CreateExcelImportForm
```

The following objects are created:

| Object | Description |
| --- | --- |
| `frmExcelImport` | Main import form |
| `frmExcelImportMap` | Field mapping subform |
| `frmExcelImportErrors` | Errors of the current batch (read-only) |
| `frmExcelImportHistory` | Import history (read-only) |
| `tblEI_Map` | Current field mapping |
| `tblEI_Batches` | Batch records |
| `tblEI_Errors` | Error details |

Running it again rebuilds and overwrites the four forms (close them first); business tables and import history are not touched. Save a copy first if you customize the forms.

Create without opening:

```vba
CreateExcelImportForm True
```

### Usage

1. Open the form:

```vba
DoCmd.OpenForm "frmExcelImport"                          ' pick the target table manually
DoCmd.OpenForm "frmExcelImport", , , , , , "Customers"   ' preset and lock the target table
```

2. Choose the target table (or preset it through `OpenArgs`).
3. Choose the Excel file (when the target table is preset, the file dialog opens automatically), then the worksheet and header row.
4. Review or adjust the "target field" of each column in the mapping grid; tick "Skip" for columns you do not want. A target field cannot be mapped twice.
5. Optionally tick "skip rows that fail validation", then click "Start import".
6. Read the completion message; open the error details and import history if needed.

### Language switch

The UI is English by default. Click the "中文 / English" button in the footer, or run:

```vba
EI_SetLanguage "zh"   ' Chinese
EI_SetLanguage "en"   ' English
```

The setting is stored in the `EI_Language` project property of the current database. Reasons already written to the log tables keep the language they were written in.

### Conversion rules

- Empty cells become Null (never 0); rows whose mapped columns are all empty are not appended.
- Text is preserved as-is. If an ID is stored in Excel as a number and padded with zeros only by display format, the leading zeros cannot be recovered; convert the source column to text first.
- Booleans: `否 / FALSE / NO / 0` → False; `是 / TRUE / YES / 1 / -1` → True.
- Dates: Excel serial numbers (1900 / 1904 date systems) and text such as `2026.10.12` are supported; other date strings follow the local regional settings, so unambiguous year-month-day formats are recommended.
- Integer fields reject fractions and overflow is reported as an error; bigint is range-checked with Decimal; Decimal is limited to the ~28-digit precision of VBA Decimal.
- Excel formula errors are reported as errors (internally marked with the reserved token `#EI_EXCEL_ERROR#`, which should not appear in normal source text).
- Unsupported fields are not imported: AutoNumber / identity, computed columns, rowversion, attachments, OLE, multi-value and GUID fields. Unmapped fields receive the database default values.

### Implementation

```
Excel worksheet
   │  Range.Value2 read in blocks (no cell-by-cell access)
   ▼
ADO client cache ── UpdateBatch ──► staging table tblEI_Stage (long-text columns)
                                       │  read in chunks via ADO Recordset; type / length / required checks in VBA memory
                                       ▼
                                  error details tblEI_Errors
                                       │  rows that pass validation
        ┌──────────────────────────────┴──────────────────────────────┐
        ▼                                                             ▼
Access target: typed staging table                       SQL Server target: batches of ≤500 rows / ≤2000 parameters
tblEI_TypedStage (UpdateBatch)                           into connection-scoped temp table #EI_Stage
        │  same connection, same transaction                           │  same connection, same transaction
        ▼                                                             ▼
INSERT INTO target SELECT … FROM staging                 INSERT INTO target SELECT … FROM #EI_Stage
```

- **Reading:** `Range.Value2` is read in blocks and only the mapped columns are used; no CSV / text driver is involved.
- **Staging:** all mapped columns are first stored as text / long text, so no type is guessed; the Excel row number is kept in a separate column to locate errors.
- **Validation:** done field by field in VBA memory. No VBA user-defined function is called from SQL, so the approach also works for external ADO / SQL Server targets.
- **Append:** Access targets are appended by one `INSERT INTO … SELECT` from a typed staging table; SQL Server targets use a direct ADO connection and append on the server from `#EI_Stage`. ADO `BeginTrans / CommitTrans / RollbackTrans` is used and the actual appended row count is verified before commit.
- **Data access:** local Access operations use `CurrentProject.Connection`; DAO is only used to read table / field definitions. Batch logging uses parameterized `ADODB.Command` with Unicode text parameters.
- **Excel instance:** only the Excel instance created by the module is closed.
- **Forms:** forms, controls and event code are all generated at runtime by `modExcelImportForms` through VBA.

### Failure handling and cancel

- Whole-batch validation by default: type / required / length errors prevent any write to the target table.
- With "skip rows that fail validation", only rows that pass validation are appended.
- Final constraints (unique index, primary key, relationships, validation rules, precision) are enforced by the database; a violation rolls back the whole append transaction and is logged as a batch-level error with `ExcelRow=0`.
- If the connection fails during commit or rollback, the batch is marked "commit status to be verified"; check the target table before retrying.
- While importing, "Close" turns into "Cancel import"; cancel is honored between chunk reads and field validations. A running SQL statement cannot be interrupted, and cancel is not possible once the commit has completed.

### SQL Server connection

By default the ODBC settings are read from the linked table's `TableDef.Connect` and used to connect directly to the back end through `Provider=MSDASQL`. The ODBC driver / DSN must match the bitness of Access, and the account needs permission to read table metadata, append data, and create temp tables in tempdb.

If the link has no saved password, or you want to use the SQL Server OLE DB driver, specify a connection at runtime:

```vba
' "Customers" is the linked table name in Access, not the SQL Server table name
EI_SetSqlServerConnection "Customers", "Provider=MSOLEDBSQL;Data Source=YOUR_SERVER;Initial Catalog=YOUR_DB;Integrated Security=SSPI;"
```

The connection string is kept only in memory for the current VBA session and is never written to source or log tables. Prefer Windows authentication and never hard-code real passwords. Passing an empty string restores the linked table's ODBC settings. After changing the connection, reselect the worksheet / target table or call `EI_Map`. Use `nvarchar` / `ntext` for Chinese text in SQL Server.

### Limitations

- Supported targets: local Access tables, linked Access back-end tables, and SQL Server user tables linked through ODBC. Other ODBC databases, SQL Server views, cross-database / four-part names, and heterogeneous links such as Excel are not supported.
- At most 254 columns can be mapped per import.
- Each user needs a separate front-end file; do not share the mapping / staging tables of one front end.
- Large imports grow the Access file; reserve disk space and compact & repair when needed.
- The project has not been verified on every combination of Access version, bitness and SQL Server driver. Test it on a backup database before use.

### License

[MIT License](LICENSE)
