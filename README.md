# excel2r

**Convert Excel formulas to executable R code.**

Upload any multi-tab `.xlsx` workbook and get a standalone `.R` script that recreates its formula logic — including cross-sheet references, conditional aggregation (SUMIFS, COUNTIFS), nested functions, and named table detection.

[![Excel2R](https://img.shields.io/badge/R-Shiny-blue)](https://img.shields.io/badge/R-Shiny-blue) [![License](https://img.shields.io/badge/license-MIT-green)](https://img.shields.io/badge/license-MIT-green)

## What This Does (and Doesn't Do)

**It does:**
- Extract every formula from your workbook and translate it to equivalent R code
- Resolve cross-sheet references and determine the correct execution order (topological sort)
- Detect named tables (ListObjects) and generate properly named R data frames with real column headers
- Produce a self-contained `.R` script that loads your Excel data and runs all calculations in R
- Flag unsupported functions clearly so nothing is silently skipped

**It doesn't:**
- Recreate the workbook visually (no formatting, charts, or cell styles)
- Replace Excel as a UI — the output is R code for scripting and automation
- Handle dynamic references (`INDIRECT`, `OFFSET`), array formulas, or structured table references (`Table1[Column]`)

**Use it when you want to:**
- **Migrate** an Excel-based workflow into R so you can extend, automate, or version-control it
- **Audit** complex workbooks by reading every formula as plain R, line by line
- **Reproduce** calculations in a scripted pipeline instead of opening Excel
- **Document** what a workbook actually computes, for handover or review

## Generated Output Explained

The `.R` script the tool produces is not a black box. Here's what each section does:

### 1. Load data — each sheet becomes an R data frame

```r
Products <- as.data.frame(openxlsx2::read_xlsx(
  excel_file, sheet = "Products",
  rows = 1:25, skip_empty_rows = FALSE, col_names = FALSE
))
colnames(Products) <- c("A", "B", "C", "D", "E")
```

Column names match Excel (A, B, C, ...) and row indices match Excel row numbers, so `Products$D[10]` in R is cell D10 in Excel.

### 2. Named tables — real column names for downstream use

When the workbook contains named tables (Insert → Table in Excel), the script generates additional data frames with proper headers:

```r
# Table "SalesData" on sheet "Products" (A1:E151)
SalesData <- Products[2:151, 1:5]
colnames(SalesData) <- c("Product", "Category", "Price", "Quantity", "Revenue")
```

The positional frames (`Products$A[10]`) are kept for formula execution. The named table frames (`SalesData$Revenue`) are for your downstream analysis — ready to use with dplyr, ggplot2, or any R workflow.

### 3. Execute formulas — in dependency order

Each formula is translated and wrapped in error handling:

```r
# D10 = SUM(D3:D9)
Products$D[10] <- tryCatch(
  sum(Products$D[3:9], na.rm=TRUE),
  error = function(e) { message('Error in Products!D10: ', e$message); NA }
)

# E5 = IF(D5>1000, D5*0.1, 0)
Products$E[5] <- tryCatch(
  ifelse(Products$D[5]>1000, Products$D[5]*0.1, 0),
  error = function(e) { message('Error in Products!E5: ', e$message); NA }
)

# Annual_Summary!B3 = 'Q1 Sales'!F20
Annual_Summary$B[3] <- tryCatch(
  Q1_Sales$F[20],
  error = function(e) { message('Error in Annual Summary!B3: ', e$message); NA }
)
```

Cross-sheet references are resolved automatically. Execution order is determined by Kahn's topological sort so dependencies are always computed first.

### 4. Verify — check what was created

```r
cat("\n=== Script execution complete ===\n")
cat("Data frames created:\n")
cat(sprintf("  Products: %d rows x %d cols\n", nrow(Products), ncol(Products)))
cat(sprintf("  Q1_Sales: %d rows x %d cols\n", nrow(Q1_Sales), ncol(Q1_Sales)))
cat(sprintf("  SalesData (table): %d rows x %d cols\n", nrow(SalesData), ncol(SalesData)))
```

After running the script, you have R data frames containing the same calculated values as your Excel workbook — ready for further analysis, plotting, or piping into other workflows.

## Features

- **62 Excel functions** mapped to R equivalents (SUM, IF, VLOOKUP, SUMIFS, INDEX/MATCH, and more)
- **Named table detection** — Excel ListObjects become properly named R data frames with real column headers
- **Auto-detects** all sheets and their actual dimensions
- **Cross-sheet references** resolved with dependency-ordered execution (Kahn's topological sort)
- **Balanced-parenthesis parser** handles nested functions like `SUM(IF(A1>0,B1,0))`
- **Downloadable .R script** — self-contained and runnable standalone
- **Interactive review** of every formula transformation before download
- **Unsupported functions** clearly flagged (not silently skipped)

## Quick Start

```r
install.packages(c("shiny", "bslib", "DT", "tidyxl", "openxlsx2", "readxl"))
shiny::runApp("path/to/excel2r-app")
```

Upload an Excel file in the browser and follow the 4-step workflow:

**Upload** → **Review** formulas → **Configure** options → **Download** .R script

## Demo

A demo workbook is included at `inst/demo/sales_report_demo.xlsx` with 5 sheets:

| Sheet | Contents |
| --- | --- |
| Products | Master product list with margins, COUNTIF |
| Q1 Sales | 19 transactions with Revenue, Net Revenue, SUMIFS by region |
| Q2 Sales | 15 transactions, same structure |
| Annual Summary | Cross-sheet refs, IFERROR, IF, SUM, AVERAGE |
| Pivot Analysis | COUNTIF, SUMIFS, nested IF (3 levels deep) |

## Supported Excel Functions

### Tested (in demo workbook or unit tests)

| Category | Functions |
| --- | --- |
| Aggregation | SUM, AVERAGE, MIN, MAX |
| Counting | COUNTIF |
| Conditional | IF (incl. 3-level nesting), IFERROR |
| Cond. Aggregation | SUMIF, SUMIFS |
| Math | ROUND, ABS |
| Logical | AND, OR, NOT |
| Text | CONCATENATE, LEFT, LEN |
| References | Cross-sheet, same-column, multi-column, whole-column, `$` absolute refs |

### Mapped but not battle-tested

| Category | Functions |
| --- | --- |
| Aggregation | MEDIAN, PRODUCT |
| Counting | COUNT, COUNTA, COUNTBLANK, COUNTIFS |
| Conditional | IFS, IFNA |
| Cond. Aggregation | AVERAGEIF, AVERAGEIFS |
| Lookup | VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP |
| Math | ROUNDUP, ROUNDDOWN, SQRT, POWER, LOG, LN, INT, MOD, EXP, SIGN |
| Text | CONCAT, RIGHT, MID, UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, VALUE |
| Info | ISNA, ISBLANK, ISNUMBER, ISTEXT, ISERROR |
| Other | ROW, COLUMN, PI |

### Not supported

`INDIRECT`, `OFFSET`, `CHOOSE`, `SWITCH`, array formulas, `TRANSPOSE`, `SORT`, `UNIQUE`, `FILTER`, `GETPIVOTDATA`, named ranges, structured table references (`Table1[Column]`).

## How It Works

```
Upload .xlsx
    │
    ▼
1. Extract formulas (tidyxl) + detect named tables (openxlsx2)
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
5. Generate self-contained .R script (with named table data frames)
    │
    ▼
Download
```

## Project Structure

```
excel2r-app/
├── app.R                        # Shiny app
├── R/                           # Core modules
│   ├── utils.R                  # Shared utilities
│   ├── extract_formulas.R       # Formula extraction via tidyxl
│   ├── detect_tables.R          # Named table detection via openxlsx2
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
setwd("path/to/excel2r-app")
source("run_tests.R")
```

Covers unit tests for each module, integration tests against the demo workbook, and shinytest2 tests for the UI.

## Dependencies

**App:** shiny, bslib, DT, tidyxl, openxlsx2, readxl

**Generated scripts use:** openxlsx2 only — conditional aggregation helpers (SUMIFS, COUNTIF, etc.) are embedded directly in the output script with no external dependencies.

## Future Work

- **Structured table references** — translate `Table1[Column]` and `[@Column]` syntax using detected table metadata
- **Named ranges** — resolve workbook-level and sheet-level named ranges into cell references

## License

MIT
