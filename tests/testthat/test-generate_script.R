# =============================================================================
# Tests for R/generate_script.R
# =============================================================================

# --- identify_criteria_columns ---

test_that("identify_criteria_columns finds SUMIFS criteria columns", {
  formula_data <- data.frame(
    Sheet = c("S1", "S1"),
    Cell = c("A1", "A2"),
    Formula = c('SUMIFS(E:E,B:B,"YES",C:C,"NO")', "SUM(D1:D10)"),
    stringsAsFactors = FALSE
  )
  result <- identify_criteria_columns(formula_data, c("S1"))
  expect_true("B" %in% result[["S1"]])
  expect_true("C" %in% result[["S1"]])
})

test_that("identify_criteria_columns finds SUMIF criteria columns", {
  formula_data <- data.frame(
    Sheet = c("S1"),
    Cell = c("A1"),
    Formula = c('SUMIF(B:B,"YES",E:E)'),
    stringsAsFactors = FALSE
  )
  result <- identify_criteria_columns(formula_data, c("S1"))
  expect_true("B" %in% result[["S1"]])
})

test_that("identify_criteria_columns returns empty for no conditional agg", {
  formula_data <- data.frame(
    Sheet = c("S1"),
    Cell = c("A1"),
    Formula = c("SUM(D1:D10)"),
    stringsAsFactors = FALSE
  )
  result <- identify_criteria_columns(formula_data, c("S1"))
  expect_equal(length(result[["S1"]]), 0)
})

# Helper to find demo file
find_demo_file <- function() {
  candidates <- c(
    "inst/demo/sales_report_demo.xlsx",
    "../../inst/demo/sales_report_demo.xlsx",
    file.path(getwd(), "inst/demo/sales_report_demo.xlsx"),
    file.path(getwd(), "../../inst/demo/sales_report_demo.xlsx")
  )
  for (f in candidates) {
    if (file.exists(f)) return(normalizePath(f))
  }
  NULL
}

# --- Full pipeline: process_excel_file ---

test_that("process_excel_file returns proper structure on demo file", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo Excel file not found")

  result <- process_excel_file(
    file_path = demo_file,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    excel_path_in_script = "sales_report_demo.xlsx"
  )

  expect_type(result, "list")
  expect_true("script" %in% names(result))
  expect_true("report" %in% names(result))
  expect_true("warnings" %in% names(result))

  # Script should be a non-empty string
  expect_true(nchar(result$script) > 100)

  # Report should be a data frame with expected columns
  expect_true("Sheet" %in% colnames(result$report))
  expect_true("Cell" %in% colnames(result$report))
  expect_true("R_Code" %in% colnames(result$report))
  expect_true("Status" %in% colnames(result$report))

  # Demo file has 5 sheets with many formulas
  expect_true(nrow(result$report) > 50)

  # Most should be ok
  ok_rate <- sum(result$report$Status == "ok") / nrow(result$report)
  expect_true(ok_rate >= 0.9, info = paste("OK rate:", ok_rate))
})

# --- Script content checks ---

test_that("generated script contains required sections", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo Excel file not found")

  result <- process_excel_file(
    file_path = demo_file,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    excel_path_in_script = "sales_report_demo.xlsx"
  )

  script <- result$script

  # Must contain key sections
  expect_true(grepl("Auto-generated R script", script))
  expect_true(grepl("library\\(openxlsx2\\)", script))
  expect_true(grepl("read_xlsx", script))
  expect_true(grepl("tryCatch", script))
  expect_true(grepl("Verification Summary", script))
})

test_that("generated script handles cross-sheet references", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo Excel file not found")

  result <- process_excel_file(
    file_path = demo_file,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    excel_path_in_script = "sales_report_demo.xlsx"
  )

  # Annual Summary sheet has cross-sheet refs to Q1 Sales and Q2 Sales
  annual_formulas <- result$report[result$report$Sheet == "Annual Summary", ]
  expect_true(nrow(annual_formulas) > 0)

  # R_Code should reference Q1_Sales and Q2_Sales data frames
  has_cross_ref <- any(grepl("Q1_Sales|Q2_Sales|Products", annual_formulas$R_Code))
  expect_true(has_cross_ref)
})

test_that("generated script handles SUMIFS formulas", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo Excel file not found")

  result <- process_excel_file(
    file_path = demo_file,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    excel_path_in_script = "sales_report_demo.xlsx"
  )

  # Should contain SUMIFS calls
  has_sumifs <- any(grepl("SUMIFS", result$report$R_Code))
  expect_true(has_sumifs)

  # Script should contain custom SUMIFS helper (no ExcelFunctionsR dependency)
  expect_true(grepl("SUMIFS <- function", result$script, fixed = TRUE))
  expect_true(grepl("\\.parse_criterion", result$script))
})

test_that("generated script handles nested IF formulas", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo Excel file not found")

  result <- process_excel_file(
    file_path = demo_file,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    excel_path_in_script = "sales_report_demo.xlsx"
  )

  # Pivot Analysis has nested IF
  has_nested_if <- any(grepl("ifelse.*ifelse", result$report$R_Code))
  expect_true(has_nested_if)
})
