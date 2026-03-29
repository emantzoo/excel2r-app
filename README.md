# excel2r

**Migrate your entire Excel workbook to R -- data and logic -- in one step.**

Upload any multi-tab `.xlsx` workbook and get a fully standalone package: tidy CSV data files and an R script (base R only, zero dependencies) that recreates every formula without needing Excel at runtime. Edit the CSVs, rerun the script, get updated results. Built-in verification compares every computed value against Excel's cached results before you commit to the migration.

[![R-CMD-check](https://github.com/emantzoo/excel2r-app/actions/workflows/R-CMD-check.yml/badge.svg)](https://github.com/emantzoo/excel2r-app/actions/workflows/R-CMD-check.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](https://opensource.org/licenses/MIT)

## Installation

Install from GitHub:

```r
# install.packages("remotes")
remotes::install_github("emantzoo/excel2r-app")
```

Then use the programmatic API:

```r
library(excel2r)

# One-step migration: script + CSVs + verification
excel2r::migrate("workbook.xlsx", "output/")

# Or process without writing files
result <- excel2r::process("workbook.xlsx")
cat(result$script)

# Verify R results against Excel
v <- excel2r::verify("workbook.xlsx", result)
print(v$summary)

# Launch the interactive Shiny app
excel2r::run_app()
```

## What This Does (and Doesn't Do)

**It does:**
- Extract every formula from your workbook and translate it to equivalent R code
- Resolve cross-sheet references and determine the correct execution order (topological sort)
- Detect named tables (ListObjects) and generate properly named R data frames with real column headers
- Export raw data as tidy CSVs — one per sheet, formula cells excluded — so the Excel file is no longer needed
- Verify computed R values against Excel's cached formula results, classifying matches, precision diffs, and real mismatches
- Produce a self-contained `.R` script using only base R — zero packages to install
- Flag unsupported functions clearly so nothing is silently skipped

**It doesn't:**
- Recreate the workbook visually (no formatting, charts, or cell styles)
- Replace Excel as a UI — the output is R code for scripting and automation
- Handle dynamic references (`INDIRECT`, `OFFSET`), array formulas, or structured table references (`Table1[Column]`)

**Use it when you want to:**
- **Migrate** — move an Excel-based workflow entirely into R, no Excel dependency at runtime
- **Automate** — plug formula logic into a pipeline, schedule it, chain it with other scripts
- **Verify** — confirm the R translation produces the same results as Excel before committing
- **Audit** — read every formula as plain R, line by line, in version-controllable code
- **Iterate** — edit input values in CSV, rerun the script, get updated results

## Two Output Modes

### Excel mode (default)
Generated script reads from the `.xlsx` file at runtime. You still need the Excel file.

### CSV standalone mode
Download a `.zip` containing:
```
excel2r_output/
├── generated_script.R    ← base R only, zero dependencies
├── data/
│   ├── Products.csv      ← tidy format: row, col, value
│   ├── Q1_Sales.csv
│   └── ...
└── README.txt
```
Delete the Excel file. Edit the CSVs to change inputs. Rerun the script. Done.

## Generated Output Explained

### 1. Data — tidy CSV per sheet

Each CSV contains only hardcoded (non-formula) cells in a compact long format:

```
row,col,value
1,A,Company Report
3,B,Revenue
3,C,Q1
4,B,North
4,C,150
5,B,South
5,C,80
7,B,Total
```

No empty cells, no formula cells. Just the raw inputs that feed into the formulas.

### 2. Grid reconstruction — CSV becomes a data frame

The generated script rebuilds the Excel grid from the tidy CSV using an embedded helper:

```r
Products <- reconstruct_grid(
  file.path(data_dir, "Products.csv"),
  max_row = 8, max_col = 4
)
```

After reconstruction, `Products$C[4]` is `150` — same position as cell C4 in Excel.

### 3. Named tables — real column names for downstream use

When the workbook contains named tables (Insert → Table in Excel), the script generates additional data frames with proper headers:

```r
# Table "SalesData" on sheet "Products" (A1:E151)
SalesData <- Products[2:151, 1:5]
colnames(SalesData) <- c("Product", "Category", "Price", "Quantity", "Revenue")
```

### 4. Formulas — executed in dependency order

Each formula is translated and wrapped in error handling:

```r
# C7 = SUM(C4:C5)
Products$C[7] <- tryCatch(
  sum(Products$C[4:5], na.rm=TRUE),
  error = function(e) { message('Error in Products!C7: ', e$message); NA }
)
```

Cross-sheet references are resolved automatically. Execution order is determined by Kahn's topological sort.

### 5. Verify — compare R results against Excel

The Verify tab runs the generated script and compares every computed value against Excel's cached formula results:

- **Exact matches** — R and Excel agree
- **Precision diffs** — minor floating-point differences (< 0.0001), classified as harmless
- **NA/error diffs** — Excel errors (#VALUE!, #REF!) or uncached formulas, classified as harmless
- **Real mismatches** — values that genuinely differ, shown with the formula, R value, Excel value, and difference

### 6. Iterate — change inputs, rerun

Change North Q1 revenue from 150 to 300 in the CSV:
```
4,"C","300"
```
Rerun the script. Totals update. No Excel involved.

## Features

- **62 Excel functions** mapped to R equivalents (SUM, IF, VLOOKUP, SUMIFS, INDEX/MATCH, and more)
- **CSV standalone mode** — export data as tidy CSVs, generated script uses base R only, zero dependencies
- **Built-in verification** — run the script and compare every result against Excel's cached values
- **Named table detection** — Excel ListObjects become properly named R data frames with real column headers
- **Auto-detects** all sheets and their actual dimensions
- **Cross-sheet references** resolved with dependency-ordered execution (Kahn's topological sort)
- **Balanced-parenthesis parser** handles nested functions like `SUM(IF(A1>0,B1,0))`
- **Interactive review** of every formula transformation before download
- **Unsupported functions** clearly flagged (not silently skipped)

## Quick Start (Shiny App)

```r
excel2r::run_app()
```

Upload an Excel file and follow the 5-step workflow:

**Upload** > **Review** formulas > **Configure** options > **Download** .R script or .zip > **Verify** against Excel

## Demo

A demo workbook is included at `inst/extdata/sales_report_demo.xlsx` with 5 sheets:

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
5. Generate output:
   ├── Excel mode: .R script (reads .xlsx at runtime)
   └── CSV mode: .zip with .R script + tidy CSVs (standalone)
    │
    ▼
6. Verify: run script, compare R values vs Excel cached results
```

## Project Structure

```
excel2r-app/
├── DESCRIPTION                  # R package metadata
├── NAMESPACE                    # Exported functions
├── R/                           # Core modules
│   ├── migrate.R                # Public API: migrate, process, verify, run_app
│   ├── excel2r-package.R        # Package-level documentation
│   ├── utils.R                  # Shared utilities
│   ├── extract_formulas.R       # Formula extraction via tidyxl
│   ├── detect_tables.R          # Named table detection via openxlsx2
│   ├── export_csv.R             # Tidy CSV export for standalone mode
│   ├── parse_formula.R          # Balanced-paren tokenizer
│   ├── transform_references.R   # Cell/range -> R syntax
│   ├── transform_functions.R    # 62 Excel functions -> R
│   ├── dependency_order.R       # Kahn's topological sort
│   ├── generate_script.R        # Script assembler (Excel + CSV modes)
│   └── verify_values.R          # R vs Excel value comparison
├── inst/app/app.R               # Shiny app (5-tab workflow)
├── inst/extdata/                # Demo Excel workbooks
├── man/                         # Auto-generated documentation
├── vignettes/                   # Getting started guide
├── tests/testthat/              # Unit & integration tests
├── .github/workflows/           # CI (R CMD check)
├── Dockerfile                   # Cloud Run deployment
└── run_tests.R                  # Dev test runner
```

## Testing

```r
# Standard R package testing
devtools::test()

# Or use the dev runner
source("run_tests.R")
```

Covers unit tests for each module (parser, transforms, dependency ordering, CSV export, verification), integration tests against the demo workbook, API tests for the public functions, and shinytest2 tests for the UI.

## Dependencies

**Package (Imports):** tidyxl, openxlsx2, readxl

**Shiny app (Suggests):** shiny, bslib, DT

**Generated scripts (Excel mode):** openxlsx2

**Generated scripts (CSV mode):** none -- base R only

## Future Work

- **Structured table references** — translate `Table1[Column]` and `[@Column]` syntax using detected table metadata
- **Named ranges** — resolve workbook-level and sheet-level named ranges into cell references

## License

MIT
