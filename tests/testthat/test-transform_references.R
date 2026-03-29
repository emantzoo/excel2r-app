# =============================================================================
# Tests for R/transform_references.R
# =============================================================================

dims <- list(
  "Sheet1" = list(max_row = 500, max_col = 20),
  "Non-Residents Tour Expenditure" = list(max_row = 100, max_col = 20)
)

# --- Single cell transforms ---

test_that("transform_single_cell handles basic references", {
  expect_equal(transform_single_cell("D10", "Sheet1"), "Sheet1$D[10]")
  expect_equal(transform_single_cell("A1", "Sheet1"), "Sheet1$A[1]")
  expect_equal(transform_single_cell("Z99", "MySheet"), "MySheet$Z[99]")
})

test_that("transform_single_cell strips dollar signs", {
  expect_equal(transform_single_cell("$D$10", "Sheet1"), "Sheet1$D[10]")
  expect_equal(transform_single_cell("D$10", "Sheet1"), "Sheet1$D[10]")
  expect_equal(transform_single_cell("$D10", "Sheet1"), "Sheet1$D[10]")
})

test_that("transform_single_cell handles cross-sheet references", {
  result <- transform_single_cell("'Non-Residents Tour Expenditure'!D9", "Sheet1")
  expect_equal(result, "NonResidents_Tour_Expenditure$D[9]")
})

test_that("transform_single_cell handles two-letter columns", {
  expect_equal(transform_single_cell("AA5", "Sheet1"), "Sheet1$AA[5]")
  expect_equal(transform_single_cell("$AB$10", "Sheet1"), "Sheet1$AB[10]")
})

# --- Range transforms ---

test_that("transform_range handles same-column ranges", {
  expect_equal(transform_range("D10:D12", "Sheet1", dims), "Sheet1$D[10:12]")
  expect_equal(transform_range("A1:A100", "Sheet1", dims), "Sheet1$A[1:100]")
})

test_that("transform_range handles whole-column ranges with real dimensions", {
  expect_equal(transform_range("A:A", "Sheet1", dims), "Sheet1$A[1:500]")
})

test_that("transform_range handles multi-column ranges", {
  result <- transform_range("A1:D10", "Sheet1", dims)
  expect_true(grepl("unlist", result))
  expect_true(grepl("Sheet1", result))
  expect_true(grepl("1:10", result))
})

test_that("transform_range handles dollar signs in ranges", {
  expect_equal(transform_range("$D$10:$D$12", "Sheet1", dims), "Sheet1$D[10:12]")
  expect_equal(transform_range("D$10:D$12", "Sheet1", dims), "Sheet1$D[10:12]")
})

test_that("transform_range handles cross-sheet ranges", {
  result <- transform_range("'Non-Residents Tour Expenditure'!D10:D20", "Sheet1", dims)
  expect_equal(result, "NonResidents_Tour_Expenditure$D[10:20]")
})

# --- Percentage transforms ---

test_that("transform_percentages converts simple percentages", {
  expect_equal(transform_percentages("100%"), "1")
  expect_equal(transform_percentages("50%"), "0.5")
  expect_equal(transform_percentages("100%-P5"), "1-P5")
})

test_that("transform_percentages preserves non-percentage content", {
  expect_equal(transform_percentages("A+B"), "A+B")
  expect_equal(transform_percentages("SUM(A1)"), "SUM(A1)")
})

test_that("transform_percentages handles multiple percentages", {
  result <- transform_percentages("100%-50%")
  expect_equal(result, "1-0.5")
})

# --- Full cell reference transform ---

test_that("transform_cell_references handles arithmetic formulas", {
  result <- transform_cell_references("D10-D11", "Sheet1", dims)
  expect_true(grepl("Sheet1\\$D\\[10\\]", result))
  expect_true(grepl("Sheet1\\$D\\[11\\]", result))
})

test_that("transform_cell_references handles SUM with range", {
  result <- transform_cell_references("SUM(D11:D12)", "Sheet1", dims)
  expect_true(grepl("Sheet1\\$D\\[11:12\\]", result))
  expect_true(grepl("SUM", result))
})

test_that("transform_cell_references handles percentages", {
  result <- transform_cell_references("100%-P5", "Sheet1", dims)
  expect_true(grepl("1-", result))
  expect_true(grepl("Sheet1\\$P\\[5\\]", result))
})
