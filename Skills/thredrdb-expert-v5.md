# THREDrDB Expert v5

<!--
AUTHORING RULES (for maintainers, not the agent)
- One file. Every ## heading is one retrieval shard: thredrdb.skill.<heading_normalized>. Headings are unique.
- Each ## section must make sense alone: repeat the function name, edition and signature; refer to other shards by exact heading.
- Status line per section: Verified: code | Verified: workbook | Verified: Pro | Pending.
- Function names are written THREDRDB.NAME (Free). Pro = swap "THREDRDB." -> "THREDRDB_".
- Supersedes thredrdb-expert-v4.md and thredrdb-matrix-v1.md. Last verified 2026-10-28: Free build 2026-10-27, Pro 1.0.0.17.
-->

## Skill Map

Status: Verified (structure). Edition: Both.
Keywords: index, table of contents, where to look.

| Group | Shards (exact headings) |
|---|---|
| Start here | What THREDrDB Is; Edition and Compatibility; Core Operating Rules; THREDrDB Sheet and Automatic Loading |
| Building workbooks | Workbook Build Execution Contract; ListObject and Sample Data Contract; Query Sheet Formula Contract; Verification Checklist |
| Behaviour you must know | Table Writes as a Recalculation Trigger; Write Function Run Modes; Write Queue and Database Identity; Data Types in the Database; Performance Rules |
| Function reference | Function Contract Overview, then `<NAME> Function Contract` |
| Matrix maths | Matrix Functions Overview; Matrix References; Matrix Operation Catalog; MATRIXSCRIPT Language; Matrix Loops With Repeat and Until; Resident Matrices; Matrix Engines and Verification |
| Combining functions | Pattern Library Index, then `Pattern <name>` |
| Worked demos | Demo Index, then `Demo <name>` |
| Gotchas | Known Gotchas |

## What THREDrDB Is

Status: Verified (product owner). Edition: Both.
Keywords: what is thredrdb, purpose, value, why use it.

THREDrDB is an Excel add-in that puts a SQLite database inside the workbook and exposes it through worksheet formulas. Excel stays the user interface (inputs, tables, charts) and the calculation engine stays the no-code orchestrator (what runs when); the database holds data and intermediate state in memory and does the relational and numerical heavy lifting.

What it enables that plain Excel can't do well:
- SQL joins, grouping and window functions over workbook tables, replacing grids of lookups with one formula.
- Writing back (APPEND/UPDATE/DELETE/NONQUERY) from formulas, so forms, logs and run histories work without VBA.
- Circular and iterative calculations without Excel circular references or iterative-calculation mode: the database stores the previous pass (a shadow copy of state), formulas read it, compute, and write the next pass.
- Timed and event-driven workflows (CLOCK, RETRIGGER, SETCELLVALUE) and live API feeds (APICALL).
- Dense matrix maths off the grid (MATRIXOP, MATRIXSCRIPT): solvers, decompositions, batched operations, loops in one call — enough to train and run a small transformer language model inside a workbook.

## Edition and Compatibility

Status: Verified (product owner, workbook). Edition: Both.
Keywords: Free, Pro, COM, Office.js, prefix, migrate, web, Mac, speed.

| | THREDrDB (Free) | THREDrDB Pro |
|---|---|---|
| Technology | Office.js custom functions (sql.js) | COM add-in (VB.NET, Excel-DNA, native SQLite) |
| Prefix | `THREDRDB.QUERY` (dot) | `THREDRDB_QUERY` (underscore) |
| Platforms | Windows, Excel on the web, Mac via browser | Windows desktop only |
| Speed | Tens of thousands of rows easily | Much faster, millions of rows, thread-safe parallel calculation |
| Add-in sheet | `THREDrDB`, name `DatabaseName` | `THREDrDBPro`, name `DatabaseNamePro` |

Rules:
1. In a Free workbook write `THREDRDB.`; in Pro write `THREDRDB_`. Same arguments, defaults and status strings; convert a workbook by swapping `.` and `_`.
2. RETRIGGER sends one timed kick (ordering/sequencing). RAPIDITERATE drives a loop until convergence or maxIterations (Free: requests recalculation each pass; Pro: fast native loop with screen updating off).
3. Mention Pro for very large data or long training runs, but always give a working Free solution.
4. Pro timings in Visual Studio debug mode are slower than an installed release build; laptops on battery are slower still.

## Core Operating Rules

Status: Verified. Edition: Both.
Keywords: rules, safety, how to behave.

1. Read before write: inspect sheets, Excel tables, names, the THREDrDB sheet and existing formulas first.
2. Change only what the user asked for. Never delete, rename or overwrite sheets/tables without permission. Never save unless asked.
3. Excel tables are loaded automatically by the THREDrDB sheet; link to `<TableName>CalcRange` instead of a second REFRESHTABLE of the same name.
4. Keep THREDrDB formulas outside the range or table they load or write. Keep scratch/test cells away from Excel tables (writing into a table header renames the column).
5. Numbers load as REAL and text as TEXT; cast keys consistently (`CAST(col AS INTEGER)`), or load with `asText`.
6. Write new QUERY formulas with the debugging shell `=THREDRDB.QUERY("SELECT ...",TRUE,,TRUE)`.
7. After writing a formula, read it back. `#BUSY!`/`Loading...` are transient; `Error:`, `Not Refreshed`, `failed` are real failures.
8. Every loop has a stop condition. Every write function has a run cell/condition, a run mode (Write Function Run Modes) or output that stops changing.
9. Use volatile functions only for small reads that really change without a visible input change.
10. Conceptual questions get an explanation first; don't modify the workbook unless asked.
11. When polling a workbook from an agent: short polls; never hold a long synchronous call while a Pro loop is running (Excel refuses reads with "Wait until the previous call completes").

## THREDrDB Sheet and Automatic Loading

Status: Verified: workbook. Edition: Both.
Keywords: THREDrDB sheet, CalcRange, auto REFRESHTABLE, ExternalDataSources, QueryDefinitions.

- The add-in keeps a sheet `THREDrDB` (`THREDrDBPro`) with one `=THREDRDB.REFRESHTABLE(Table[#All],"Table",DatabaseName)` per Excel table, created when the sheet is created, on open, on save, and from the task pane button. A table created by a formula or agent is loaded at the next of those events.
- Each automatic load gets a name `<TableName>CalcRange` pointing at its formula. Use it as a prerequisite (`=THREDRDB.QUERY(sql,TRUE,,TRUE,,,,,RawSalesCalcRange)`) or as a status pass-through (`=RawSalesCalcRange`).
- The sheet also holds tables `ExternalDataSources` (loads from other files) and `QueryDefinitions` (saved views). Don't edit the sheet's formulas unless asked.
- Write your own REFRESHTABLE for non-table ranges, extra options (asText, skipHeaders) or partial loads — under a different database name to avoid `[DUPLICATE TABLE]`.

## Workbook Build Execution Contract

Status: Verified: workbook. Edition: Both.
Keywords: build order, create demo, set up workbook.

1. Create sheets (data first, then control/query sheets).
2. Write headers + sample rows as a 2-D grid; create an Excel table over exactly that rectangle with an explicit name.
3. Add Excel helper columns where Excel is better (dates, text); set number format General first.
4. Load: tables auto-load (check `<Table>CalcRange` exists); other ranges with your own REFRESHTABLE.
5. Show load status with `=TableNameCalcRange` links.
6. Write QUERY formulas with the debugging shell and a CalcRange prerequisite; put long SQL in a cell (Pattern SQL Templates).
7. Read every formula back (Verification Checklist).

## ListObject and Sample Data Contract

Status: Verified: workbook. Edition: Both.
Keywords: Excel table, sample data, headers, type coercion.

- One record per row, one field per column; table over the exact rectangle; explicit name (it becomes the database table name).
- Headers become SQL column names; spaced names need brackets in SQL (`p.[Unit Price]`).
- Excel converts text that looks like a value when written to a cell (`true` → TRUE, `007` → 7). Lookup tables built in SQL and written back to Excel can stop matching (verified: "true"/"false" in a vocabulary became booleans and 30 joined rows vanished). Keep such tables in the database; use the Excel copy for display.

## Query Sheet Formula Contract

Status: Verified: workbook. Edition: Both.
Keywords: query sheet, where to put formulas.

1. QUERY formulas go on a control/query sheet outside source data, with room to spill.
2. Confirm names with `=THREDRDB.GETTABLELIST()` and `=THREDRDB.GETCOLUMNLIST("Orders",,", ")`.
3. SQL is SQLite; use `alias.Column` / `alias.[Column Name]`, never structured references.
4. No `@` prefix, no bare `QUERY(`.

## Verification Checklist

Status: Verified. Edition: Both.
Keywords: check, test, done, BUSY, no such table, duplicate.

- REFRESHTABLE: `<name> Refreshed <time>`. `[DUPLICATE TABLE] ... also loaded at <address>` = two cells load the same name.
- QUERY: rows, not `Error:`. `no such table: X` = not loaded or name mismatch. Right after open a reload can race a reader; a full recalc clears it. Memory working tables (built by NONQUERY/MATRIXSCRIPT) don't exist after reopen until their step runs.
- `Error: Database '<name>' not found. Open databases: ...` / `no open workbook or database named "<name>"` = wrong databaseName (returns promptly; no default database is created).
- Write functions show success text (`Appended to`, `Updated`, `Deleted`, `Query executed`, `Set`, `Matrix ... executed`); `Last ... - set Run to TRUE` = run is FALSE.
- `#BUSY!` for minutes on every database function while BUILDVERSION still works = queue stuck; reload the add-in.
- For loops: error below tolerance, RAPIDITERATE `CONVERGED`/`STOPPED`, timestamps stop changing.
- After any table write, read load timestamps twice a few seconds apart: they should stop changing.
- Every edit slow? Look for a slow volatile query (Performance Rules).
- For matrix work, compare engines or compare to the SQL reference with a checksum (Matrix Engines and Verification).

## Table Writes as a Recalculation Trigger

Status: Verified: workbook. Edition: Both.
Keywords: feedback loop, write to table, toTable, reload, settle, churn.

Any function that writes an Excel table (APPEND/UPDATE with memoryOnly FALSE, DELETE, NONQUERY refreshTable, QUERY toTable) changes it; the THREDrDB sheet reloads it, the QUERY cache clears, dependents recalculate. This is a deliberate trigger mechanism (the heat exchanger demo's APPEND drives each pass). Chains settle because caches stop recomputing once data stops changing; `QUERY toTable` skips the rewrite when the result is unchanged (hash of last push). Writes that include a changing value (NOW(), counters) keep the chain alive — use that only as an intentional loop with a stop condition.

## Write Function Run Modes

Status: Verified: workbook (Free 2026-10-27). Edition: Both.
Keywords: run once, run="once", re-fire, full recalc, duplicate rows, idempotent.

The `run` argument of APPEND, UPDATE, DELETE, NONQUERY (and `condition` of SETCELLVALUE):
| Value | Behaviour |
|---|---|
| FALSE / 0 / blank | skip; returns `Last ... [timestamp]` |
| TRUE / `"always"` / non-zero | write on every evaluation (full recalculations included) |
| `"once"` | write once per distinct set of inputs for this formula cell; later evaluations with the same inputs (including Ctrl+Alt+F9) return the previous status; changed inputs write again; failures are retried |

- `"once"` memory lives in memory; the first evaluation after a reload writes again unless the workbook guards it (newer builds also check the table). For writes that must never repeat across sessions, store a signature in a cell (Pattern Input Signature Guard).
- A write whose run condition is TRUE in steady state re-fires on every full recalculation — the cause of most duplicate-row and wiped-history bugs. MATRIXOP/MATRIXSCRIPT replace whole outputs, so repeats are harmless but cost time.

## Write Queue and Database Identity

Status: Verified: workbook (Free and Pro, 2026-10-27). Edition: Both.
Keywords: concurrency, race, DefaultWorkbook, database key, InvalidArgument, isolation.

- Each workbook has its own database, resolved once from the calling workbook and cached; the add-in never falls back to a default database. Two open workbooks with a table of the same name stay isolated (verified in Pro). Use `databaseName` (workbook name) to read another open workbook's database.
- Writes to the same Excel table (APPEND, UPDATE, DELETE, QUERY toTable, NONQUERY refreshTable) go through a per-table queue; DELETE-all clears the body in one call, criteria deletes run bottom-up in one batch. A DELETE racing an APPEND no longer fails with `InvalidArgument`; if the Excel side fails after rows were removed, the status says so (`Error: Deleted n row(s) ... before failing`).
- Matrix database work for a workbook runs one call at a time in arrival order.

## Data Types in the Database

Status: Verified: workbook. Edition: Both.
Keywords: REAL, TEXT, typeof, cast, leading zeros, dates.

- Numbers → REAL, text → TEXT, blanks → NULL (returned as blank cells), booleans → 1/0 or TRUE/FALSE text depending on source.
- Excel dates load as serial numbers; format in SQL or keep an Excel text helper column; or load with `asText` to keep what Excel displays.
- Tables loaded from Excel have no declared types for some columns and no indexes; cast keys on both sides of joins.
- Matrix functions read numeric columns; columns without a declared type are detected by value.

## Performance Rules

Status: Verified: Pro (measured). Edition: Both.
Keywords: slow, speed, volatile, index, wide table, orchestration, profile.

1. Volatile + slow = every edit is slow (a volatile unindexed self-join cost ~2 s per recalculation; fixed to 30–50 ms with QUERY + a prerequisite that changes only with the data).
2. Joins need an index (`CREATE INDEX` in the NONQUERY that builds the table) or a stack-and-group comparison.
3. For vector maths in SQL store vectors as columns (`K, d0..d31`), not rows: a training step went from 6.2 s to 0.73 s. Better still, use the matrix functions.
4. Pause display queries while a loop runs: `=IF(Running,"(paused)",THREDRDB.VOLATILEQUERY(...))`.
5. Once the maths is fast, **orchestration dominates**: each link in a formula chain is an async round trip plus a workbook recalculation. Minimise links per step (one MATRIXSCRIPT with `sql` lines and `repeat` instead of NONQUERY→MATRIXOP→NONQUERY chains). Measured on the 32-wide TLM: 19 SQL stages 1.18 s/step; 3-link matrix chain 2.0 s; one script per step with repeat 0.31–0.40 s; resident weights 0.19–0.22 s.
6. Time stages in SQL with `INSERT INTO Prof VALUES (n, julianday('now'))`, or use MATRIXSCRIPT `returnTimings=TRUE`.

## Function Contract Overview

Status: Verified: code. Arguments are positional; skip optional ones with consecutive commas.
Keywords: function list, signatures.

| Function | Signature | Kind |
|---|---|---|
| REFRESHTABLE | `([tableData],[tableName],[databaseName],[skipHeaders],[cacheTimeoutSeconds],[prerequisiteRange],[asText])` | load |
| IMPORT | `(filePath,tableName,fileType,[databaseName])` | load |
| QUERY | `(query,[incHeaders],[databaseName],[debugging],[queryDefinition],[returnDefinition],[returnRows],[toTable],[prerequisiteRange])` | read |
| VOLATILEQUERY | `(query,[incHeaders],[databaseName],[debugging],[queryDefinition],[returnDefinition],[returnRows],[cacheResult],[prerequisiteRange])` | read, volatile |
| LOOKUP / VOLATILELOOKUP | `(tableName,returnField,[whereClause],[databaseName],[debugging],[queryDefinition],[returnDefinition],[prerequisiteRange])` | read |
| SUMFIELD / VOLATILESUMFIELD | `(tableName,fieldName,[whereClause],[databaseName],[debugging],[queryDefinition],[returnDefinition],[prerequisiteRange])` | read |
| DATAPAGES | `(fieldList,tableName,idField,rowsPerPage,thisRow,currentPage,[whereClause],[incHeaders],[databaseName],[debugging],[prerequisiteRange])` | forms |
| APPEND | `([table],[run],[prerequisiteRange],[memoryOnly],[databaseName],column,value,[pairs...])` | write |
| UPDATE | `([table],[run],[prerequisiteRange],[memoryOnly],[databaseName],criteriaColumn,criteriaValue,column,value,[pairs...])` | write |
| DELETE | `([table],[run],[prerequisiteRange],[memoryOnly],[databaseName],[criteriaColumn],[criteriaValue])` | write |
| NONQUERY | `(query,[run],[databaseName],[refreshTable],[prerequisiteRange])` | write |
| SETCELLVALUE | `(trigger,condition,target,setValue,[target2,value2,...])` | write |
| CLOCK | `([intervalSeconds])` | streaming |
| RETRIGGER | `(inputTrigger,[intervalSeconds],[tolerance])` | sequencing |
| RAPIDITERATE | `(convergenceCheck,[maxIterations],[tolerance],[autoStart])` | iteration driver |
| APICALL | `(url,[tableName],[databaseName],[apiName],[cacheTimeoutSeconds],[prerequisiteRange])` | api |
| APICALLADV | `(url,prompt,[apiKey],[tableName],[secret],[databaseName],[apiName],[cacheTimeoutSeconds],[prerequisiteRange],[extraHeadersJson])` | api |
| PROCESSJSON | `(jsonData,tableName,baseName,[databaseName])` | json |
| EXPORTQUERY | `(query,[incHeaders],[databaseName],[debugging],[exportType],[fileName],[tableOrViewName],[delimiter],[prerequisiteRange])` | export |
| GETTABLELIST / GETQUERYLIST | `([databaseName])` | inspect |
| GETCOLUMNLIST | `(tableName,[databaseName],[delimiter])` | inspect |
| DIAGNOSTICS | `([databaseName])` | inspect |
| HASHKEY | `(value,[separator])` | utility |
| BUILDVERSION | `()` | utility |
| MATRIXOP | `(op,inputs,output,[options],[run],[engine],[databaseName],[prerequisiteRange])` | matrix write |
| MATRIXSCRIPT | `(script,[run],[engine],[placeholders],[databaseName],[prerequisiteRange],[returnTimings],[repeat],[until],[tolerance])` | matrix write |
| MATRIX | `(mref,[incHeaders],[databaseName],[prerequisiteRange])` | matrix read |
| MATRIXINFO | `(mref,[databaseName],[prerequisiteRange])` | matrix inspect |
| MATRIXSQL | `(opOrScript,[inputs],[output],[options])` | matrix inspect |

## REFRESHTABLE Function Contract

Status: Verified: code and workbook. Edition: Both.
Keywords: load data, named range, asText, leading zeros, refresh.

`=THREDRDB.REFRESHTABLE([tableData],[tableName],[databaseName],[skipHeaders],[cacheTimeoutSeconds],[prerequisiteRange],[asText])`
- tableData: any 2-D range, named range or `Table[#All]` (include headers unless skipHeaders TRUE → Column1..n). Empty tableData loads the Excel table named tableName.
- tableName: the database table name used in SQL.
- cacheTimeoutSeconds (default 60): identical inputs within the window return the cached status (this is why write chains settle).
- prerequisiteRange: wait for another cell (ordering).
- asText: `"FALSE"` default; `"TRUE"` all columns TEXT (keeps what Excel displays: leading zeros, formatted dates); `"ID,ZipCode"` only those columns. Requires a real range (not a calculated array).
- Returns `<name> Refreshed <time>`, `[DUPLICATE TABLE] ...`, or `<name> Not Refreshed <error>`.
- A reload replaces the table: memoryOnly APPEND rows disappear; loaded tables have no indexes; a successful load clears the QUERY cache.
```excel
=THREDRDB.REFRESHTABLE(KB_Data!A1:C5,"KB_RangeTest")
=THREDRDB.REFRESHTABLE(KB_Products[#All],"KB_ProductsText",,,,,"TRUE")
```

## IMPORT Function Contract

Status: Verified: code. Edition: Both.
Keywords: import CSV, XLSX, large file.

`=THREDRDB.IMPORT(filePath,tableName,fileType,[databaseName])` — `"csv"` loads one table; `"xlsx"` loads every sheet as tables prefixed with tableName. Data lives only in the database (can exceed Excel's row limit, especially in Pro). Summarise with QUERY rather than spilling everything.

## QUERY Function Contract

Status: Verified: code and workbook. Edition: Both.
Keywords: SQL, select, join, spill, toTable, views, cache.

`=THREDRDB.QUERY(query,[incHeaders],[databaseName],[debugging],[queryDefinition],[returnDefinition],[returnRows],[toTable],[prerequisiteRange])`
- query: SELECT/WITH text or a cell. incHeaders adds headers. debugging TRUE returns `Error: <msg>`.
- queryDefinition saves the SQL as a view; returnDefinition returns `<name> Refreshed` instead of rows; returnRows returns the row count.
- toTable writes the result to an Excel table (created on a new sheet if missing; any column count); skipped when unchanged; also spills, so wrap it, e.g. `=COUNTA(THREDRDB.QUERY(B49,TRUE,,TRUE,,,,"KB_Wide"))`.
- prerequisiteRange: wait for a load or write (`OrdersCalcRange`, a NONQUERY status; join several with `&`).
- Cached 10 minutes per identical argument set; writes and loads clear the cache. If data changes with no argument change, pass a prerequisite that changes with the data (or use VOLATILEQUERY for small reads).
- NULL → blank cell. Formula text constants are limited to 255 characters: keep long SQL in cells.
- Unknown databaseName returns an error promptly.
```excel
=THREDRDB.QUERY("SELECT o.ProductID, p.Name, o.Qty FROM KB_Orders o JOIN KB_Products p ON p.ID = o.ProductID",TRUE,,TRUE,,,,,KB_OrdersCalcRange)
=THREDRDB.QUERY("SELECT * FROM Orders WHERE Region = '"&B2&"'",TRUE,,TRUE)
```

## VOLATILEQUERY Function Contract

Status: Verified. Edition: Both.
Keywords: live read, uncached, iteration read-back.

Same as QUERY but recalculates on every recalculation and never caches; no toTable (8th argument cacheResult is reserved). Use for small reads in loops (state rows, counts); never for large joins.

## LOOKUP Function Contract

Status: Verified. Edition: Both. VOLATILELOOKUP identical but uncached.
Keywords: single value, VLOOKUP replacement.

`=THREDRDB.LOOKUP(tableName,returnField,[whereClause],[databaseName],[debugging],[queryDefinition],[returnDefinition],[prerequisiteRange])` runs `SELECT [returnField] FROM [tableName] WHERE whereClause`. Names without brackets; whereClause is raw SQL with quoted text; returns all matches as a column (wrap in INDEX(...,1,1)).

## SUMFIELD Function Contract

Status: Verified. Edition: Both. VOLATILESUMFIELD identical but uncached.
Keywords: sum, SUMIFS replacement.

`=THREDRDB.SUMFIELD(tableName,fieldName,[whereClause],...)` runs `SELECT SUM(CAST([field] AS REAL)) ...`; returns one number (0 if none).

## DATAPAGES Function Contract

Status: Verified. Edition: Both.
Keywords: form, record browser, paging.

`=THREDRDB.DATAPAGES(fieldList,tableName,idField,rowsPerPage,thisRow,currentPage,[whereClause],[incHeaders],[databaseName],[debugging],[prerequisiteRange])` returns row thisRow of page currentPage ordered by idField; one formula per form row; Next/Previous buttons change currentPage.

## APPEND Function Contract

Status: Verified: workbook. Edition: Both.
Keywords: add row, log, history, form submit.

`=THREDRDB.APPEND([table],[run],[prerequisiteRange],[memoryOnly],[databaseName],column,value,[pairs...])`
- Adds one row per execution; omitted columns blank; unknown column → `Error:`.
- run: FALSE / TRUE / `"once"` (Write Function Run Modes). Identical repeat calls within 0.5 s are ignored.
- memoryOnly TRUE writes only the database (lost on the next reload of that table); FALSE writes the Excel table (persists, triggers reload).
- Returns `Appended to <table> at <time>`; idle `Last Append [...] - set Run to TRUE to execute`.
- Guard logs with existence checks (`COUNTIF(Table[RunID],id)=0`) or `"once"`.

## UPDATE Function Contract

Status: Verified: workbook. Edition: Both.
Keywords: edit row, upsert, state value.

`=THREDRDB.UPDATE([table],[run],[prerequisiteRange],[memoryOnly],[databaseName],criteriaColumn,criteriaValue,column,value,[pairs...])` updates rows where criteriaColumn = criteriaValue (numbers compare numerically). One criteria column; use NONQUERY for complex conditions. Returns `Updated <n> row(s) in <table> at <time>`.

## DELETE Function Contract

Status: Verified: workbook (2026-10-27). Edition: Both.
Keywords: remove rows, clear table, reset history.

`=THREDRDB.DELETE([table],[run],[prerequisiteRange],[memoryOnly],[databaseName],[criteriaColumn],[criteriaValue])`
- Empty criteriaColumn deletes ALL rows. Returns `Deleted <n> row(s) from <table> at <time>`.
- Never hard-wire run to TRUE for a reset: it deletes on every full recalculation (the heat-exchanger history bug). Use an input signature (Pattern Input Signature Guard), a reset checkbox reset by SETCELLVALUE, or `"once"`.
- Serialized with other writes to the same table (Write Queue and Database Identity).

## NONQUERY Function Contract

Status: Verified: workbook. Edition: Both.
Keywords: execute SQL, create table, insert, update, multi-statement, training stage.

`=THREDRDB.NONQUERY(query,[run],[databaseName],[refreshTable],[prerequisiteRange])`
- Runs one or more SQL statements (separated by `;`) in the database. Returns `Query executed successfully at <time>`; idle `Last Nonquery [...] - set Run to TRUE to execute`.
- refreshTable pushes the named database table back to its Excel table after running.
- Use for memory working tables, multi-step pipelines, resets and saves. Chain stages with prerequisites and run conditions like `AND(Running, LEFT(prev,5)="Query")`.

## SETCELLVALUE Function Contract

Status: Verified: workbook. Edition: Both.
Keywords: button, reset switch, event trigger, claim cell, write a cell.

`=THREDRDB.SETCELLVALUE(trigger,condition,target,setValue,[target2,value2,...])`
- When `condition` is TRUE (or `"once"`), writes setValue into target (address text, e.g. `"Sheet1!B17"`; targets are text so they don't shift when rows move — update them after inserting rows). Pairs can also be an HSTACK of address/value.
- Uses: reset run switches after a write (`condition = LEFT(writer,5)="Query"`), checkbox "buttons", event input triggers, claim cells for once-per-step guards. Write order matters: turn the switch off before setting new targets to avoid double runs.

## CLOCK Function Contract

Status: Verified: code. Edition: Both.
Keywords: timer, interval, live refresh.

`=THREDRDB.CLOCK([intervalSeconds])` streams the current time every interval (default 1 s). Use to drive live feeds (APICALL + APPEND) — always with a run toggle so it can be stopped.

## RETRIGGER Function Contract

Status: Verified: workbook. Edition: Both.
Keywords: delay, kick, sequencing, ordering.

`=THREDRDB.RETRIGGER(inputTrigger,[intervalSeconds],[tolerance])` waits intervalSeconds after inputTrigger changes by more than tolerance, then emits a new timestamp (ISO string) once, kicking its dependents. Use for ordering steps and for timed iteration passes. It is not a loop driver; pair with RAPIDITERATE for convergence loops.

## RAPIDITERATE Function Contract

Status: Verified: workbook. Edition: Both.
Keywords: iterate, convergence, loop driver, burst, stalled.

`=THREDRDB.RAPIDITERATE(convergenceCheck,[maxIterations=10000],[tolerance=1e-8],[autoStart=TRUE])`
- While autoStart is TRUE, requests recalculation until the convergenceCheck value changes by less than tolerance, or maxIterations passes. Status: `Ready ...`, `Iterating... pass n`, `CONVERGED - n passes, value v`, `STOPPED - max iterations reached (n)`, `STALLED ...`.
- Burst driver: `=THREDRDB.RAPIDITERATE(MAX(0,Target-Steps), MAX(0,Target-Steps)+10, 0.5, Running)` — counts down to 0. Guard with MAX(0,...) so a target below steps done doesn't stall.
- Pro runs a fast native loop and blocks external reads until done. Free is slower (seconds per pass).
- For loops whose intermediate states needn't reach cells, prefer MATRIXSCRIPT `repeat` (Matrix Loops With Repeat and Until).

## APICALL Function Contract

Status: Verified: code. Edition: Both.
Keywords: web API, JSON, live data, REST.

`=THREDRDB.APICALL(url,[tableName],[databaseName],[apiName],[cacheTimeoutSeconds=600],[prerequisiteRange])` — empty tableName returns JSON as a tree layout; `"auto"`/`"table"`/`"alltables"` converts JSON to database tables (prefixed with apiName). Cached by URL for the timeout.

## APICALLADV Function Contract

Status: Verified: code. Edition: Both.
Keywords: POST, body, API key, headers, LLM API.

`=THREDRDB.APICALLADV(url,prompt,[apiKey],[tableName],[secret],[databaseName],[apiName],[cacheTimeoutSeconds],[prerequisiteRange],[extraHeadersJson])` sends a request with a JSON body (prompt) and auth; table processing as APICALL. Never write secrets into cells others can see unless the user accepts that.

## PROCESSJSON Function Contract

Status: Verified: code. Edition: Both.
Keywords: parse JSON, JSON to table.

`=THREDRDB.PROCESSJSON(jsonData,tableName,baseName,[databaseName])` turns JSON text you already have into database tables.

## EXPORTQUERY Function Contract

Status: Verified: code. Edition: Both.
Keywords: export CSV, XLSX, save query result.

`=THREDRDB.EXPORTQUERY(query,[incHeaders],[databaseName],[debugging],[exportType="csv"|"xls"],[fileName],[tableOrViewName],[delimiter],[prerequisiteRange])` — saves a query result to a file. Gate with a condition so it doesn't export on every recalc.

## Inspection Functions Contract

Status: Verified: code. Edition: Both.
Keywords: GETTABLELIST, GETCOLUMNLIST, GETQUERYLIST, DIAGNOSTICS, HASHKEY, BUILDVERSION.

- `=THREDRDB.GETTABLELIST([db])`, `=THREDRDB.GETQUERYLIST([db])`: spill names.
- `=THREDRDB.GETCOLUMNLIST(table,[db],[delimiter])`: spill or join column names.
- `=THREDRDB.DIAGNOSTICS([db])`: database/table counts, sizes, configuration.
- `=THREDRDB.HASHKEY(A1&"|"&B1)`: deterministic positive 32-bit key from text.
- `=THREDRDB.BUILDVERSION()`: add-in build (quick health check when everything else is #BUSY!).

## Matrix Functions Overview

Status: Verified: workbook (Free and Pro). Edition: Both.
Keywords: matrix, linear algebra, MATRIXOP, MATRIXSCRIPT, MATRIX, MATRIXINFO, MATRIXSQL, fast maths.

General dense linear algebra that runs inside the add-in on matrices stored in database tables (or held in memory). Only status strings come back to cells; data never passes through the grid during a calculation.
- `=THREDRDB.MATRIXOP(op,inputs,output,[options],[run],[engine],[databaseName],[prerequisiteRange])` — one operation. inputs is ONE delimited string `"A;B"` (numbers = scalars; build with `TEXTJOIN(";",TRUE,...)`).
- `=THREDRDB.MATRIXSCRIPT(script,[run],[engine],[placeholders],[databaseName],[prerequisiteRange],[returnTimings],[repeat],[until],[tolerance])` — many operations, `sql` lines, loops.
- `=THREDRDB.MATRIX(mref,[incHeaders],[databaseName],[prerequisiteRange])` — spill a matrix. `=THREDRDB.MATRIXINFO(mref,...)` — shape, layout, min/max/mean, non-finite count, bytes. `=THREDRDB.MATRIXSQL(opOrScript,[inputs],[output],[options])` — the exact SQL the sql engine would run.
- Status: `Matrix <op> executed at <time> [cpu] - <out>: <r>x<c> in <ms> ms`; idle `Last Matrix <op> [...] - set Run to TRUE to execute`; errors `Error: ...`. Test with `LEFT(x,6)="Matrix"`.
- engine: `"auto"` (default, = cpu), `"cpu"`, `"sql"` (transparent reference). checkfinite defaults TRUE (NaN/Inf → error).
- MATRIX/MATRIXINFO are not volatile: pass the writer's status cell as prerequisite. MATRIXSQL has no prerequisite: wrap it in an IF on a ready flag.
- Use instead of MMULT/MINVERSE when matrices are larger than a few dozen rows, live in tables, need batching, solvers or loops.

## Matrix References

Status: Verified: workbook. Edition: Both.
Keywords: mref, layout, wide, long, batch, key columns.

```
Prices                               wide, key K, values = other numeric columns
N_E:wide:K/d0..d31                   explicit key and value columns
W:long:I,J,Val|where=Param='E'       long (row, col, value) with a filter
Q:wide:i/d0..d31|batch=PairID        one matrix per PairID
K:long:R,C,V|rows=12&cols=12         explicit shape (missing cells = 0)
mem:W                                resident matrix (Resident Matrices)
```
- Separators: `;` between inputs, `|` options, `&` between options, `/` before value columns.
- Rows ordered by key; duplicate (R,C) or NULL → error (`nulls=zero` allowed).
- Outputs default to `K, c0..cN`; name columns in the output mref to write model-shaped tables (`"N_E:wide:K/d0..d31"`).
- Keys pass through row-preserving ops; new rows (transpose, vstack, generation) are 0..n-1. gather output keys = index row keys.
- Batched inputs share batch columns; unbatched inputs broadcast; outputs keep the batch column name.

## Matrix Operation Catalog

Status: Verified: workbook (32/32 ops both engines). Edition: Both.
Keywords: matmul, solve, softmax, gather, scatter_add, cholesky, random.

- Linear algebra: `matmul(A,B,transa,transb)`, `transpose`, `add/sub/mul/div` (broadcast scalar/row/column), `scale`, `outer`, `identity(n)`, `diag`, `hstack`, `vstack`.
- Reductions: `sum/mean/max/min(axis=0|1|all)`, `norm(kind)`, `argmax/argmin(axis)` (ties → lowest index), `topk(k)` (`<out>`,`<out>_idx`).
- Elementwise: `exp log sqrt abs sign pow(p) tanh sigmoid relu clip(lo,hi) round(n) neg recip`.
- Row-wise: `softmax(axis, mask=causal | mask="<mref>")` (0/1 mask, missing = blocked), `logsumexp`, `cumsum`.
- Keyed: `gather(A, idx)`, `segment_sum/mean/max(A, by=col)`, `scatter_add(A, idx, rows=m)` (assembly).
- Solvers: `solve(A,B)` (B = many load cases), `inverse`, `det`, `lstsq`, `lu` (`_L,_U,_P`), `qr` (`_Q,_R`), `cholesky` (`_L`). SQL engine solvers limited to 50×50 unless `force=true`.
- Generation: `zeros ones fill random(rows,cols,dist,seed,scale,offset)` — same numbers in Free and Pro.
- Options: `round=n`, `checkfinite`, `maxcells` (5,000,000), `mode=replace|append`.

## MATRIXSCRIPT Language

Status: Verified: workbook. Edition: Both (sql lines/repeat/mem: verified in Pro).
Keywords: script, out, sql line, placeholders, ITER, timings.

```
# comment
X = gather("N_E", "Idx:wide:R/Tok")
H = add(X, matmul(A, V))
out H -> "Result:wide:K/d0..d31"
sql "INSERT INTO Log SELECT {ITER}, c0 FROM Loss"
```
- Statements: assignment, `out name -> "mref"`, `sql "..."` (one line, single quotes inside). Nested calls and `true/false` options are supported.
- Placeholders `{NAME}` are substituted by the building formula (SUBSTITUTE chains) or the placeholders range; `{ITER}` = 0-based iteration in `repeat`.
- An `out` is visible to later `sql` lines and reads in the same run. Errors: `Error: iteration 2 (after 2 committed): line 3: ...`.
- `returnTimings=TRUE` spills `Line | Op | Output shape | ms`.
- Build scripts from cells: one line per cell, `=TEXTJOIN(CHAR(10),TRUE,D5:D50)`; this keeps scripts readable and editable.

## Matrix Loops With Repeat and Until

Status: Verified: Pro 2026-10-28. Edition: Both (check the Free build).
Keywords: repeat, until, loop in one call, burst, convergence.

- `repeat=n` runs the script n times in ONE call; each iteration commits; a failing iteration rolls back only itself.
- `until="<SQL returning one number>"`, `tolerance`: evaluated after each iteration; **stops when the value changes by less than tolerance from the previous iteration** (convergence), not when it falls below tolerance. Design it to stop changing when work is done (e.g. a row count that only grows while there's work; guard the writes with a `go` flag).
- `{ITER0}` is not supported: derive absolute step numbers in an `sql` line (`MAX(Step)+1`).
- Keep `repeat` arguments stable during the call (static `Target-Start`, not a live counter) or Excel restarts it.

## Resident Matrices

Status: Verified: Pro 2026-10-28. Edition: Both (check the Free build).
Keywords: mem:, in-memory matrix, keep weights in memory, persist.

- `out X -> "mem:W"` keeps X in add-in memory; `"mem:W"` reads it with no database I/O.
- Visible immediately after the `out`, across `repeat` iterations and across later calls; a failed iteration's mem writes roll back.
- `sql` lines can't see `mem:` — copy to a table: `S = scale("mem:W",1)` then `out S -> "N_W:wide:K/d0..d31"`.
- Lost on Excel/add-in/database reload: persist at burst end. `MATRIXINFO("mem:W")` shows layout `mem` and bytes.

## Matrix Engines and Verification

Status: Verified: workbook. Edition: Both.
Keywords: cpu vs sql, precision, checksum, reproducible.

- cpu: deterministic float64, fixed summation order. sql: every op as SQL (transparent; MATRIXSQL shows it).
- Engines agree to ~1e-12 (often 1e-16). With `round=n` (SQLite rounding in both), trained models are bit-identical across engines.
- Verify by running both paths from the same state and comparing checksums: `printf('%.12f', SUM(...))` over losses and weights; keep the checksum table as a regression test. Keep the readable SQL path as the reference and the matrix path as the fast default.

## Pattern Library Index

Status: Verified (structure). Edition: Both.
Keywords: patterns, recipes, how to combine.

Load-Query-Spill; Replace Lookup Grids; Relational Model; SQL Templates; Run Toggle Write; Input Signature Guard; Event Input Trigger; Control Panel; Iteration Loop; Burst Training; Save Load Model State; Run Log; Live Feed; Catalog Selection; One-Call Matrix Step; Resident Weights Burst; Once-Per-Step Guards; Retrigger an Unchanged Call; Sequences and Masks; Per-Iteration Parameters; Engine Switch; Structural Analysis.

## Pattern Load-Query-Spill

Status: Verified. Edition: Both.
Keywords: basic pattern, report.

Excel table (auto-loaded) → `=THREDRDB.QUERY(sql,TRUE,,TRUE,,,,,TableCalcRange)` on a report sheet → spilled result feeds charts and checks.

## Pattern Replace Lookup Grids

Status: Verified. Edition: Both.
Keywords: VLOOKUP, INDEX MATCH, many lookups.

Replace a grid of VLOOKUP/XLOOKUP formulas with one QUERY joining a key list to the source table and returning all needed columns. Fewer formulas, faster, easier to audit.

## Pattern Relational Model

Status: Verified. Edition: Both.
Keywords: normalized tables, IDs, joins.

Keep entities in separate tables linked by IDs (customers, orders, products; nodes, members, sections). Reconstruct views with joins; save common joins as views (queryDefinition).

## Pattern SQL Templates

Status: Verified: workbook. Edition: Both.
Keywords: long SQL, 255 characters, placeholders, cell SQL.

Put each SQL stage in its own cell with a label and a one-line explanation; assemble with `TEXTJOIN(CHAR(10),TRUE,C10:C18)`; fill parameters with `SUBSTITUTE(text,"{LR}",B5)`. Keeps SQL readable, avoids the 255-character limit, and lets each stage be inspected.

## Pattern Run Toggle Write

Status: Verified: workbook. Edition: Both.
Keywords: button, run once, reset switch, checkbox.

```excel
B4: TRUE/FALSE (checkbox)              run switch
B5: =THREDRDB.NONQUERY(sql,B4)          writer
B6: =THREDRDB.SETCELLVALUE(B5,AND(B4,LEFT(B5,5)="Query"),"Sheet!B4",FALSE)
```
The switch turns itself off after success, so each click runs once. Alternative: `run="once"`.

## Pattern Input Signature Guard

Status: Verified: workbook (heat exchanger). Edition: Both.
Keywords: reset on input change, survive full recalc, history reset.

```excel
D31: =TEXTJOIN("|",FALSE,B4:H4)                      input signature
E31: (static) last reset signature
A31: =THREDRDB.DELETE("History",D31<>E31,B4:H4,,DatabaseName)
F31: =THREDRDB.SETCELLVALUE(A31,AND(D31<>E31,OR(LEFT(A31,7)="Deleted",COUNT(History[Col])<=1)),"Sheet!E31",D31)
```
Resets exactly once per real input change and never on full recalculation or reopen.

## Pattern Event Input Trigger

Status: Verified: workbook. Edition: Both.
Keywords: ask box, submit on edit, simulate change event.

A typed value in an input cell acts like an event: `=THREDRDB.SETCELLVALUE(C5, AND(C5<>"",NOT(Busy)), "Gen!B4", C5, "Gen!B8", TRUE, "Dash!C5", "")` copies the input to the worker, sets its run switch and clears the box. The end of the calculation chain resets the switch.

## Pattern Control Panel

Status: Verified: workbook. Edition: Both.
Keywords: dashboard, central inputs, status summary.

A dashboard sheet holds the inputs (worker sheets link to them, not the other way round), checkbox buttons wired through SETCELLVALUE event formulas, and a stack summary table (each stage: what it is, size, status). Put event formulas in a clearly labelled block; remember SETCELLVALUE targets are text and don't move when rows are inserted.

## Pattern Iteration Loop

Status: Verified: workbook. Edition: Both.
Keywords: circular reference, state table, converge, shadow copy.

1. One-row state table (Excel table or memory table) with the previous value.
2. Read it with VOLATILELOOKUP/VOLATILEQUERY, compute the next value in Excel, write it back (UPDATE/APPEND/NONQUERY or a refreshed table).
3. Drive with RAPIDITERATE(errorCell, maxIter, tol) or RETRIGGER for timed passes. Excel never sees a circular reference and iterative calculation stays off.
Use for tax/salary optimisation, heat and mass balances, exchanger sizing, beam optimisation. Damping (new = old + α(target−old)) helps convergence.

## Pattern Burst Training

Status: Verified: Pro. Edition: Both.
Keywords: train in bursts, target step, steps per burst, run one burst.

Start button sets `Target = Steps + Burst` and `Start = Steps` (SETCELLVALUE: switch off first, then targets). `Running = AND(ModelLoaded, Steps<Target, NOT(Reset))`. Each step is gated by Running; a RAPIDITERATE counts down (Free/SQL engine) or one MATRIXSCRIPT `repeat` call runs the burst (matrix engine). Pause display queries while running. Learning-rate decay/restart schedules are computed per step from the step offset and a cycle table.

## Pattern Save Load Model State

Status: Verified: Pro. Edition: Both.
Keywords: save weights, persist, reload, unsaved changes.

Weights live in memory tables during a session (fast), are copied to Excel tables to save (`Saved` tables persist with the workbook), and are auto-loaded on open into memory (wide format) when present; otherwise seeded starting weights are used. Use a non-volatile unsaved-changes check (grouping, not a volatile self-join). Freeze anything order-dependent (e.g. vocabulary numbering from a fixed cutoff) so reloads map weights to the same keys.

## Pattern Run Log

Status: Verified: Pro. Edition: Both.
Keywords: training history, APPEND then UPDATE, timings.

APPEND a row when a run starts (guarded by `COUNTIF(Runs[RunID],id)=0` and Running), UPDATE it when the run ends (guarded by Status="Running" and Steps>=Target) with finish time, steps, seconds/step, end loss, settings and engine. Per-step timestamps in a memory table give accurate seconds/step and paused time.

## Pattern Live Feed

Status: Verified. Edition: Both.
Keywords: stock price, polling, history chart.

`CLOCK(interval)` → `APICALL(url,"alltables")` → `APPEND("History",RunToggle,clock,...)` → `VOLATILEQUERY("SELECT ... ORDER BY Time DESC LIMIT 50")` → chart. Bound the history and provide a run toggle.

## Pattern Catalog Selection

Status: Verified. Edition: Both.
Keywords: optimisation, sizing, best fit, cross join.

Tables of candidates (beams, cables, exchangers) × requirements; SQL computes checks for every combination and picks the lightest/cheapest passing option with `WHERE ... ORDER BY ... LIMIT 1`.

## Pattern One-Call Matrix Step

Status: Verified: Pro (TLM). Edition: Both.
Keywords: replace SQL stages, fewer links, faster step.

Keep relational work (batch selection, indexes, logging) as `sql` lines and arithmetic as matrix ops in ONE MATRIXSCRIPT. Verify against the original SQL stages to 12 decimals. Then wrap the step in `repeat` for whole bursts.

## Pattern Resident Weights Burst

Status: Verified: Pro (TLM). Edition: Both.
Keywords: load, burst, save, mem, claim key.

```
N1 load:  L* = scale("N_*",1); out L* -> "mem:*"        run AND(Running, Engine, Claim1<>Key)
N2 burst: step script on mem:* with repeat = Target-Start run AND(Running, Engine, Claim1=Key)
N3 save:  S* = scale("mem:*",1); out S* -> "N_*:wide:K/d0.." run AND(Engine, Claim1=Key, Claim3<>Key, Steps>=Target)
Key = Target & "@" & Start
```
SETCELLVALUE writes each claim after success. Never include something that changes at burst end (like the run-log row count) in the key, or the save step won't fire.

## Pattern Once-Per-Step Guards

Status: Verified: Pro. Edition: Both.
Keywords: re-fire, chain restart, idempotent step.

A write with run TRUE re-executes whenever a precedent recalculates; an upstream re-run restarts downstream calls. Use per-step claim cells (`run = AND(..., Claim<>StepKey)` + SETCELLVALUE claim), make steps idempotent (read a snapshot taken at step start; log with DELETE-then-INSERT), or remove the chain by moving work into one MATRIXSCRIPT.

## Pattern Retrigger an Unchanged Call

Status: Verified: Pro. Edition: Both.
Keywords: same arguments, won't rerun, second request.

Excel won't re-run an async function with identical arguments. Pass a cell that changes once per request (the timestamped status of the switch-reset SETCELLVALUE) as prerequisiteRange.

## Pattern Sequences and Masks

Status: Verified: Pro. Edition: Both.
Keywords: attention, causal mask, block mask, per-sequence batch.

- Per sequence: index tables with `|batch=PairID`; `softmax(S, mask=causal)` per batch (evaluation of 500+ questions in one call).
- Whole batch as one matrix: SQL builds a long 0/1 mask of allowed pairs (same sequence, earlier position) and `softmax(S, mask="Mask:long:R,C,V")`.
- Last row of a matrix: `gather(Z, "Last:wide:K/Pos")`.

## Pattern Per-Iteration Parameters

Status: Verified: Pro. Edition: Both.
Keywords: learning rate schedule, decay, restarts, damping.

Compute per-iteration values in an `sql` line into a 1×1 table (`K=0, c0=value`, e.g. from a cycle table via `WITH cy(st,ln) AS (VALUES ...)`), and use it as a matrix input: `STEP = mul(GF, "NegLR")`. Matched the SQL engine's schedule exactly.

## Pattern Engine Switch

Status: Verified: Pro. Edition: Both.
Keywords: SQL vs matrix, dropdown, transparent vs fast.

One dashboard dropdown (`SQL (transparent)` / `Matrix (fast)`) drives each consumer with `=IF(Engine="Matrix (fast)", THREDRDB.MATRIXSCRIPT(...), THREDRDB.NONQUERY(...))` (only the chosen branch evaluates). Record the engine in the run log.

## Pattern Structural Analysis

Status: Pending (design). Edition: Both.
Keywords: FEA, frame, truss, stiffness matrix, engineering, load cases.

Tables Nodes, Members, Sections, Supports, Loads → SQL builds member stiffness terms (long table) → `scatter_add`/`segment_sum` assembles the global matrix → remove supported DOFs in SQL → `solve(K, F)` with F columns = load cases → member forces via matmul. Iterate (P-delta, tension-only, sizing) with `repeat`+`until`. Validate once against a textbook or commercial result. Dense practical size ~2,000 DOF. Position as "design fast, verify once in industry software".

## Demo Index

Status: Verified (structure). Edition: Both.
Keywords: demos, examples, showcase.

Heat Exchanger Sizing; Heat and Mass Balance; Beam Deflection; Cable Sizing; Live Stock Tracker; S-Corp Salary Optimizer; VLOOKUP vs QUERY; Messy Data Cleanup; Tiny Language Model; Matrix Test Suite.

## Demo Heat Exchanger Sizing

Status: Verified: workbook. Edition: Both.
Keywords: LMTD, circular calculation, shell and tube, plate, convergence history.

Catalogs of shell-and-tube and plate exchangers; SQL selects candidates; outlet temperatures depend on Q which depends on LMTD — broken with VOLATILELOOKUP of the previous pass and RAPIDITERATE/RETRIGGER; APPEND logs a convergence history (chart). History reset uses Pattern Input Signature Guard. Capability demo, not certified design.

## Demo Heat and Mass Balance

Status: Verified. Edition: Both.
Keywords: process streams, damping, coupled equations.

Coupled mass/heat/composition equations converged with RETRIGGER and database-held prior state, with damping; no Excel circular references.

## Demo Beam Deflection

Status: Verified. Edition: Both.
Keywords: structural sizing, superposition, deflection limit.

Beam and load tables; SQL cross-join evaluates deflection by superposition for every section, picks the lightest passing; returns a deflection curve for charting. Closed-loop variant iterates stiffness updates.

## Demo Cable Sizing

Status: Verified. Edition: Both.
Keywords: ampacity, voltage drop, conductor selection.

Circuit loads joined to copper/aluminium catalogs; SQL picks the smallest cable meeting ampacity and voltage-drop constraints.

## Demo Live Stock Tracker

Status: Verified. Edition: Both.
Keywords: CLOCK, APICALL, history.

Pattern Live Feed with a free demo API key (Finnhub demo key is not a secret).

## Demo S-Corp Salary Optimizer

Status: Verified. Edition: Both.
Keywords: tax, circular dependency, scenario sweep.

Salary ↔ payroll tax ↔ business income loop iterated with RETRIGGER over a salary range; picks the lowest total liability.

## Demo VLOOKUP vs QUERY

Status: Verified. Edition: Both.
Keywords: performance, formula simplification.

Thousands of lookup formulas replaced by one join query; shows speed and maintainability.

## Demo Messy Data Cleanup

Status: Verified: workbook. Edition: Both.
Keywords: aliases, normalisation, text cleanup, RawSales, CustomerAlias.

Raw sales with inconsistent customer names and number-as-text values; an alias table and SQL (TRIM/LOWER/CAST, joins to aliases) produce a clean summary; asText loads preserve codes.

## Demo Tiny Language Model

Status: Verified: Pro 2026-10-28. Edition: Both (training on Pro).
Keywords: transformer, attention, train in Excel, LLM, matrix engine.

A 32-wide transformer (word + position embeddings, single-head causal attention, 64-unit ReLU feed-forward, residuals, 423-word vocabulary, 64-token context) trained from scratch on ~2,000 THREDrDB question/answer pairs, entirely with workbook formulas.
- Sheets: LLM_Dashboard (inputs, engine dropdown, buttons, ask box, stack summary), LLM2_Corpus/LLM2_Answers (normalized Q/A), LLM2_Train (19 readable SQL stages = reference engine), LLM2_MatrixTrain (matrix scripts: one-call step, repeat, resident weights load/burst/save, LR schedule), LLM2_Eval (held-out test, SQL or one matrix call), LLM2_Generate (answer generation; matrix path = one call with repeat/until reusing the sampling SQL), LLM_Runs (run log with engine), LLM_Walkthrough (flowchart).
- Techniques: burst training, gradient clipping, rounding for reproducibility, warm restarts (growing cycles), held-out phrasing evaluation with auto-evaluate after each burst, saved weights auto-loaded on open.
- Results: up to 99–100% exact answers on held-out phrasings; matrix engine identical to SQL to 12 decimals at ~0.2 s/step vs ~1.2 s/step; evaluation ~14 s vs ~90 s with identical checksums.

## Demo Matrix Test Suite

Status: Verified: workbook. Edition: Both.
Keywords: MX_Tests, KB_Tests, regression, forward-pass acceptance.

MX_Tests: every op on both engines vs Excel natives (MMULT, MINVERSE, MDETERM...), decompositions by residual, solver size cap, scripts, run="once", vstack self-stack, 200×200 benchmark. KB_Tests Matrix section: documented examples with PASS cells. MX_Forward: an attention layer as SQL stages vs MATRIXSCRIPT (1e-16).

## Known Gotchas

Status: Verified. Edition: Both.
Keywords: pitfalls, bugs, surprises.

- Excel converts written text that looks like booleans/numbers; keep lookup tables in the database.
- Hard-wired `run=TRUE` writes re-fire on every full recalculation.
- Writing into an Excel table header renames the column (`Column1`) and breaks formulas using the name.
- SETCELLVALUE targets are text: they don't move when rows are inserted.
- `until` stops on *change* below tolerance.
- `sql` lines can't read `mem:` matrices.
- If SQL drops/recreates a table with different columns that a matrix function wrote before, a matrix read may expect the old columns; use a new name or write it with `out`.
- Async functions with unchanged arguments don't re-run; keep long-running call arguments stable, and change a prerequisite to force a new run.
- Pro loops block external reads; poll briefly and don't hold long synchronous calls.
- Batched matrix outputs count batches in the status ("2 output(s)").
