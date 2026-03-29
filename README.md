# excel2r

**Convert Excel formulas to executable R code.**

Upload any multi-tab `.xlsx` workbook and get a standalone `.R` script that recreates its formula logic — including cross-sheet references, conditional aggregation (SUMIFS, COUNTIFS), nested functions, and more.

![Excel2R](https://img.shields.io/badge/R-Shiny-blue) ![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **62 Excel functions** supported (SUM, IF, VLOOKUP, SUMIFS, INDEX/MATCH, ROUND, and many more)
- **Auto-detects** all sheets and their actual dimensions — no hardcoded limits
- **Cross-sheet references** resolved with dependency-ordered execution (Kahn's topological sort)
- **Balanced-parenthesis parser** handles nested functions like `SUM(IF(A1>0,B1,0))`
- **Downloadable .R script** that is self-contained and runnable standalone
- **Interactive review** of every formula transformation before download
- **Unsupported functions** are clearly flagged (not silently skipped)

## Quick Start

```r
# Install dependencies
install.packages(c("shiny", "bslib", "DT", "tidyxl", "openxlsx2", "readxl"))

# Run the app
shiny::runApp("path/to/excel2r-app")
```

Then upload an Excel file in the browser and follow the 4-step workflow:
1. **Upload** → 2. **Review** formulas → 3. **Configure** options → 4. **Download** .R script

## Demo

A demo workbook is included at `inst/demo/sales_report_demo.xlsx` with 5 sheets:

| Sheet | Contents |
|-------|----------|
| Products | Master product list with margins, COUNTIF |
| Q1 Sales | 19 transactions with Revenue, Net Revenue, SUMIFS by region |
| Q2 Sales | 15 transactions, same structure |
| Annual Summary | Cross-sheet refs, IFERROR, IF, SUM, AVERAGE |
| Pivot Analysis | COUNTIF, SUMIFS, nested IF (3 levels deep) |

## Excel Functions: Tested vs Mapped

The tool has transformation rules for 62 Excel functions, but **not all have been battle-tested with real-world workbooks**. The table below shows the honest status:

### Tested with real formulas (included in demo workbook or unit tests)

| Category | Functions | Test Coverage |
|----------|-----------|---------------|
| **Aggregation** | SUM, AVERAGE, MIN, MAX | Tested in demo + unit tests |
| **Counting** | COUNTIF | Tested in demo workbook |
| **Conditional** | IF (incl. 3-level nesting), IFERROR | Tested in demo workbook |
| **Cond. Aggregation** | SUMIF, SUMIFS | Tested in demo + real-world workbook |
| **Math** | ROUND, ABS | Tested in unit tests |
| **Logical** | AND, OR, NOT | Tested in unit tests |
| **Text** | CONCATENATE, LEFT, LEN | Tested in unit tests |
| **References** | Cross-sheet refs, same-column ranges, multi-column ranges, whole-column ranges, `$` absolute refs | All tested |

### Mapped but not battle-tested (transform rules exist, may have edge cases)

| Category | Functions | Notes |
|----------|-----------|-------|
| **Aggregation** | MEDIAN, PRODUCT | Simple wrappers, likely fine |
| **Counting** | COUNT, COUNTA, COUNTBLANK, COUNTIFS | COUNTIFS delegates to ExcelFunctionsR |
| **Conditional** | IFS, IFNA | IFS uses `dplyr::case_when()` — may need tuning |
| **Cond. Aggregation** | AVERAGEIF, AVERAGEIFS | Delegates to ExcelFunctionsR |
| **Lookup** | VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP | Custom helpers generated — approximate match may not be exact |
| **Math** | ROUNDUP, ROUNDDOWN, SQRT, POWER, LOG, LN, INT, MOD, EXP, SIGN | Simple wrappers |
| **Text** | CONCAT, RIGHT, MID, UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, VALUE | Simple wrappers |
| **Info** | ISNA, ISBLANK, ISNUMBER, ISTEXT, ISERROR | Simple wrappers |
| **Other** | ROW, COLUMN, PI | Limited context support |

### Not supported (will be flagged as unsupported in output)

- `INDIRECT`, `OFFSET` — dynamic references, cannot be statically translated
- `CHOOSE`, `SWITCH` — not yet implemented
- Array formulas (`{=SUM(IF(...))}` with Ctrl+Shift+Enter) — no special handling
- `TRANSPOSE`, `SORT`, `UNIQUE`, `FILTER` — dynamic array functions
- `GETPIVOTDATA` — pivot table references
- Named ranges — not resolved (tidyxl extracts formulas with the names, not their definitions)
- Structured table references (`Table1[Column]`) — not parsed

## How It Works

```
Upload .xlsx
    │
    ▼
1. Extract formulas (tidyxl)
    │
    ▼
2. Tokenize & parse (balanced-paren parser)
    │
    ▼
3. Transform: cell refs → R syntax, functions → R equivalents
    │
    ▼
4. Determine execution order (Kahn's topological sort)
    │
    ▼
5. Generate self-contained .R script
    │
    ▼
Download
```

## Generated Script Structure

The output `.R` script contains:

1. **Package setup** — `openxlsx2`, `ExcelFunctionsR` (only if needed)
2. **Helper functions** — VLOOKUP, MATCH, ROUNDUP equivalents (only if used)
3. **Data loading** — reads each sheet into an R data frame with proper types
4. **Formula execution** — all formulas in dependency order, wrapped in `tryCatch()`
5. **Verification summary** — prints data frame dimensions

## Project Structure

```
excel2r-app/
├── app.R                        # Shiny app
├── R/                           # Core modules (auto-sourced by Shiny)
│   ├── utils.R                  # Shared utilities
│   ├── extract_formulas.R       # Formula extraction via tidyxl
│   ├── parse_formula.R          # Balanced-paren tokenizer
│   ├── transform_references.R   # Cell/range → R syntax
│   ├── transform_functions.R    # 62 Excel functions → R
│   ├── dependency_order.R       # Kahn's topological sort
│   └── generate_script.R        # Script assembler
├── tests/testthat/              # 150+ unit & integration tests
├── inst/demo/                   # Demo Excel workbook
└── run_tests.R                  # Test runner
```

## Testing

```r
# Run all tests
setwd("path/to/excel2r-app")
source("run_tests.R")
```

Tests cover:
- **Unit tests** for each module (utils, parser, transforms, dependency ordering)
- **Integration tests** against the demo workbook
- **Shinytest2** tests for the web app UI

## Dependencies

**App runtime:**
- `shiny`, `bslib`, `DT` — web UI
- `tidyxl` — formula extraction
- `openxlsx2` — data reading
- `readxl` — sheet listing

**Generated scripts use:**
- `openxlsx2` — data reading
- `ExcelFunctionsR` — SUMIFS/COUNTIFS (only when needed)

## License

MIT
