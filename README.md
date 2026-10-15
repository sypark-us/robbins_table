# OfficeTable

A spreadsheet that runs entirely in the browser: an Excel-style grid with a real formula engine,
charts, sorting and filtering, conditional formatting, data validation, and `.xlsx` / `.csv`
import and export. No server, no install, no build step; your data never leaves your computer.

**Run it:** serve the folder and open `index.html` (browsers block ES modules on `file://`):

```
npm start                      # node test/serve.mjs 8080, then open http://localhost:8080/
```

Any static file server works (`python -m http.server`, VS Code Live Server, GitHub Pages).

Tests (Node 20+; the browser test needs Chrome or Edge installed):

```
npm run test:unit        # node --test: formula engine, functions, model, undo, formats, CSV/XML/ZIP/XLSX
npm run test:smoke       # drives the served app in headless Chrome: editing, clipboard, fill, undo,
                         # formatting, conditional formats, sort/filter, charts, dialogs, file round trips
npm test                 # both
node test/screenshot.mjs screenshots   # renders a sample workbook and saves PNGs
```

## Features

### Grid
- 1,048,576 rows × 16,384 columns, virtualised canvas rendering with frozen panes, hidden rows and
  columns, merged cells, zoom (25–400 %), column/row resize and autofit (double-click the header edge).
- Excel keyboard model: arrows, Shift/Ctrl+arrows, Tab/Enter, Ctrl+Home/End, Page Up/Down,
  Ctrl+Space / Shift+Space, Ctrl+A (region, then sheet), F2, Esc, Delete, Alt+Enter, Ctrl+Enter,
  Ctrl+D / Ctrl+R, Ctrl+; and Ctrl+Shift+;, Ctrl+1, Ctrl+B/I/U/5, Ctrl+Shift+1/2/3/4/5/~, Ctrl+9/0,
  Ctrl+Plus/Minus, Ctrl+PgUp/PgDn, Shift+F11, F9, Alt+=, Ctrl+` and more (Help → Keyboard shortcuts).
- Text overflow into empty neighbours, wrap text, indent, horizontal and vertical alignment,
  number display that shrinks digits or shows `####` when a number does not fit.

### Formulas
- A1 references (relative/absolute, `$A$1`), ranges, whole rows and columns (`A:A`, `3:5`),
  sheet-qualified references (`'Data 2'!B2`), defined names, array constants, `%`, `&`, `^`.
- 300+ worksheet functions: math and trigonometry, statistics (AVERAGEIFS, MEDIAN, STDEV, PERCENTILE,
  RANK, CORREL, NORM.DIST…), text (TEXT, TEXTJOIN, TEXTSPLIT, SUBSTITUTE…), logical (IF, IFS, SWITCH,
  LET, IFERROR), lookup (VLOOKUP, HLOOKUP, XLOOKUP, INDEX/MATCH, XMATCH, OFFSET, INDIRECT, CHOOSE),
  dynamic arrays (FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, TRANSPOSE, HSTACK, VSTACK, TAKE, DROP…),
  dates (DATEDIF, EDATE, EOMONTH, NETWORKDAYS, WORKDAY, YEARFRAC…), financial (PMT, PV, FV, NPV, IRR,
  XIRR, RATE, IPMT, PPMT, SLN, DDB…), engineering (DEC2BIN, BITAND, CONVERT…), database (DSUM…),
  SUBTOTAL/AGGREGATE that respect filtered rows. Help → Function list shows every signature.
- Dynamic arrays spill into neighbouring cells (`#SPILL!` when blocked), array arithmetic broadcasts.
- Incremental recalculation through a dependency graph, circular-reference detection (`#CIRC!`),
  volatile functions (NOW, TODAY, RAND, INDIRECT…), Excel error values with hover explanations.
- Formula bar and in-cell editing with function autocomplete, argument tooltips, coloured
  reference highlighting, click-and-drag reference picking, F4 to cycle `$`.
- References adjust when rows/columns are inserted or deleted, cells are cut and pasted, or sheets
  are renamed; copied formulas shift relative references.

### Formatting
- Fonts, size, bold/italic/underline/strikethrough, text and fill colour, borders (presets and per
  side styles), alignment, wrap, indent, merge.
- Excel number-format codes: `#,##0.00`, `0%`, `$#,##0.00;[Red]-$#,##0.00`, accounting, scientific,
  fractions, dates and times (`yyyy-mm-dd`, `[h]:mm:ss`, `dddd, mmmm d`), conditional sections and
  colours, custom codes. Typed input such as `1,234.5`, `50%`, `$12`, `2024-01-05`, `Jan 5`, `3:45 PM`
  is recognised and formatted automatically.
- Conditional formatting: cell value, text contains/begins/ends, duplicate/unique, top/bottom N,
  above/below average, blanks, custom formula, 2- and 3-colour scales, data bars.
- Format painter, Paste special (values, formats, formulas, transpose, skip blanks).

### Data
- Sort (quick A→Z / Z→A on the current region, multi-level custom sort with header detection).
- AutoFilter with per-column value lists, text and number conditions; filtered rows are skipped by
  copy and by SUBTOTAL(1xx).
- Data validation (list with in-cell dropdown, whole/decimal/date, text length, custom formula) with
  stop / warning / information behaviour and input prompts.
- Remove duplicates, text to columns, trim whitespace, autofill series (numbers, dates by day or
  month, weekday and month names, `Item 1` patterns, formulas), Ctrl-drag to force increments.
- Named ranges (name box or Name manager), notes with hover display, hyperlinks (Ctrl+click).

### Charts
- Clustered and stacked column/bar, line (with markers), area, stacked area, pie, doughnut, scatter.
- Series/category detection from the selected range, titles, legends, data labels, hover tooltips,
  drag to move, corner handle to resize, right-click to change type or export as PNG. Charts follow
  the data they reference.

### Files
- Open `.xlsx` (shared strings, shared formulas, styles, merges, column widths, frozen panes,
  conditional formats, validations, hyperlinks, defined names, date1904), `.csv` / `.tsv` (delimiter
  sniffing) and OfficeTable `.json`. Drag and drop onto the window works too.
- Save as `.xlsx` (uses the File System Access API when available, otherwise downloads), export the
  current sheet as CSV/TSV, save as JSON to keep charts. Print the sheet or selection.
- The workbook is kept in the browser's local storage between visits (File → toggle).

## Layout

```
index.html        app shell
css/style.css
js/app.js         commands, keyboard, selection, files, sheets, status bar
js/grid.js        canvas renderer, hit testing, scrolling, mouse interaction
js/editor.js      in-cell editor / formula bar, autocomplete, reference picking
js/menus.js       menu bar, toolbar, context menus
js/dialogs.js     modal dialogs and popovers
js/clipboard.js   copy/cut/paste, TSV and HTML interchange
js/fill.js        autofill series detection
js/charts.js      chart data extraction, canvas rendering, floating chart objects
js/model.js       workbook, sheets, cells, styles, axes, undo/redo, structural edits
js/engine.js      dependency graph, recalculation, spilling
js/formula.js     tokenizer, parser, reference rewriting
js/functions.js   worksheet function library
js/format.js      number-format engine, date serials, input parsing
js/xlsx.js        XLSX import/export      js/zip.js  ZIP      js/xml.js  XML      js/csv.js  CSV
test/             node --test suites, browser smoke test, screenshot script
```

## Known limits

- Charts, notes and format-painter state are not written to `.xlsx` (save as JSON to keep charts).
- Pivot tables, macros, sparklines, R1C1 notation and 3-D references (`Sheet1:Sheet3!A1`) are not supported.
- `#CIRC!` is shown for circular references instead of Excel's iterative calculation.

MIT License.
