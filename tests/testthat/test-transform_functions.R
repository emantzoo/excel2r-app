# =============================================================================
# Tests for R/transform_functions.R
# =============================================================================

# --- Registry checks ---

test_that("registry has all expected functions", {
  supported <- get_supported_functions()
  expect_true(length(supported) >= 50)

  # Core functions must be present
  core_funcs <- c("SUM", "AVERAGE", "MIN", "MAX", "IF", "IFERROR",
                   "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
                   "VLOOKUP", "INDEX", "MATCH",
                   "ROUND", "ABS", "SQRT", "POWER", "LOG",
                   "CONCATENATE", "LEFT", "RIGHT", "MID", "LEN",
                   "AND", "OR", "NOT",
                   "ISNA", "ISBLANK")
  for (fn in core_funcs) {
    expect_true(fn %in% supported, info = paste(fn, "should be supported"))
  }
})

test_that("is_function_supported works", {
  expect_true(is_function_supported("SUM"))
  expect_true(is_function_supported("sum"))  # case insensitive
  expect_false(is_function_supported("INDIRECT"))
  expect_false(is_function_supported("OFFSET"))
})

# --- Aggregation functions ---

test_that("SUM transforms correctly", {
  result <- transform_function_call("SUM", c("Sheet1$A[1:10]"))
  expect_equal(result$status, "ok")
  expect_equal(result$code, "sum(Sheet1$A[1:10], na.rm=TRUE)")
})

test_that("SUM with multiple args uses c()", {
  result <- transform_function_call("SUM", c("A[1]", "A[2]", "A[3]"))
  expect_equal(result$status, "ok")
  expect_true(grepl("sum\\(c\\(", result$code))
})

test_that("AVERAGE transforms correctly", {
  result <- transform_function_call("AVERAGE", c("Sheet1$A[1:10]"))
  expect_equal(result$code, "mean(Sheet1$A[1:10], na.rm=TRUE)")
})

test_that("MIN and MAX transform correctly", {
  r_min <- transform_function_call("MIN", c("A[1:10]"))
  r_max <- transform_function_call("MAX", c("A[1:10]"))
  expect_true(grepl("^min\\(", r_min$code))
  expect_true(grepl("^max\\(", r_max$code))
})

# --- Conditional functions ---

test_that("IF transforms to ifelse", {
  result <- transform_function_call("IF", c("A[1]>0", "B[1]", "0"))
  expect_equal(result$code, "ifelse(A[1]>0, B[1], 0)")
})

test_that("IF with 2 args defaults false value", {
  result <- transform_function_call("IF", c("A[1]>0", "B[1]"))
  expect_equal(result$code, "ifelse(A[1]>0, B[1], FALSE)")
})

test_that("IFERROR transforms to tryCatch", {
  result <- transform_function_call("IFERROR", c("1/0", "NA"))
  expect_true(grepl("tryCatch", result$code))
})

test_that("IFNA transforms correctly", {
  result <- transform_function_call("IFNA", c("A[1]", "0"))
  expect_true(grepl("is.na", result$code))
})

# --- Conditional aggregation ---

test_that("SUMIFS passes through to ExcelFunctionsR", {
  result <- transform_function_call("SUMIFS", c("A[1:10]", "B[1:10]", '"YES"'))
  expect_equal(result$code, 'SUMIFS(A[1:10], B[1:10], "YES")')
})

test_that("COUNTIF passes through", {
  result <- transform_function_call("COUNTIF", c("A[1:10]", '">5"'))
  expect_true(grepl("COUNTIF", result$code))
})

# --- Lookup functions ---

test_that("VLOOKUP generates helper call", {
  result <- transform_function_call("VLOOKUP", c("x", "table", "3", "FALSE"))
  expect_equal(result$code, "excel_vlookup(x, table, 3, FALSE)")
})

test_that("INDEX generates helper call", {
  result <- transform_function_call("INDEX", c("A[1:10]", "5", "1"))
  expect_true(grepl("excel_index", result$code))
})

test_that("MATCH generates helper call", {
  result <- transform_function_call("MATCH", c("x", "A[1:10]", "0"))
  expect_true(grepl("excel_match", result$code))
})

# --- Math functions ---

test_that("ROUND transforms correctly", {
  result <- transform_function_call("ROUND", c("3.14159", "2"))
  expect_equal(result$code, "round(3.14159, 2)")
})

test_that("ABS transforms correctly", {
  result <- transform_function_call("ABS", c("-5"))
  expect_equal(result$code, "abs(-5)")
})

test_that("POWER transforms to ^ operator", {
  result <- transform_function_call("POWER", c("2", "3"))
  expect_equal(result$code, "(2) ^ (3)")
})

test_that("LOG with base transforms correctly", {
  result <- transform_function_call("LOG", c("100", "10"))
  expect_equal(result$code, "log(100, base = 10)")
})

test_that("LN transforms to natural log", {
  result <- transform_function_call("LN", c("2.718"))
  expect_equal(result$code, "log(2.718)")
})

# --- Text functions ---

test_that("CONCATENATE transforms to paste0", {
  result <- transform_function_call("CONCATENATE", c('"hello"', '" "', '"world"'))
  expect_true(grepl("paste0", result$code))
})

test_that("LEFT transforms to substr", {
  result <- transform_function_call("LEFT", c('"hello"', "3"))
  expect_true(grepl("substr", result$code))
})

test_that("LEN transforms to nchar", {
  result <- transform_function_call("LEN", c('"hello"'))
  expect_equal(result$code, 'nchar("hello")')
})

# --- Logical functions ---

test_that("AND transforms to & operator", {
  result <- transform_function_call("AND", c("A>0", "B>0"))
  expect_equal(result$code, "(A>0 & B>0)")
})

test_that("OR transforms to | operator", {
  result <- transform_function_call("OR", c("A>0", "B>0"))
  expect_equal(result$code, "(A>0 | B>0)")
})

test_that("NOT transforms to ! operator", {
  result <- transform_function_call("NOT", c("A>0"))
  expect_equal(result$code, "!(A>0)")
})

# --- Unsupported functions ---

test_that("unsupported functions return proper status", {
  result <- transform_function_call("INDIRECT", c('"A1"'))
  expect_equal(result$status, "unsupported")
  expect_true(grepl("Unsupported", result$code))
})

# --- Full pipeline: transform_all_functions ---

test_that("transform_all_functions handles nested SUM(IF(...))", {
  # After reference transform, this might look like:
  formula <- "SUM(IF(Sheet1$A[1]>0,Sheet1$B[1],0),Sheet1$C[1])"
  result <- transform_all_functions(formula)
  expect_true(grepl("sum", result$formula))
  expect_true(grepl("ifelse", result$formula))
  expect_equal(length(result$warnings), 0)
})

test_that("transform_all_functions handles formula with no functions", {
  result <- transform_all_functions("Sheet1$A[1]+Sheet1$B[1]")
  expect_equal(result$formula, "Sheet1$A[1]+Sheet1$B[1]")
})

test_that("transform_all_functions collects warnings for unsupported", {
  result <- transform_all_functions("INDIRECT(Sheet1$A[1])")
  expect_true(length(result$warnings) > 0)
})
