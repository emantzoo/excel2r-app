# =============================================================================
# Tests for R/utils.R
# =============================================================================

test_that("col_letter_to_index converts correctly", {
  expect_equal(col_letter_to_index("A"), 1)
  expect_equal(col_letter_to_index("B"), 2)
  expect_equal(col_letter_to_index("Z"), 26)
  expect_equal(col_letter_to_index("AA"), 27)
  expect_equal(col_letter_to_index("AB"), 28)
  expect_equal(col_letter_to_index("AZ"), 52)
  expect_equal(col_letter_to_index("BA"), 53)
  expect_equal(col_letter_to_index("ZZ"), 702)
})

test_that("index_to_col_letter converts correctly", {
  expect_equal(index_to_col_letter(1), "A")
  expect_equal(index_to_col_letter(26), "Z")
  expect_equal(index_to_col_letter(27), "AA")
  expect_equal(index_to_col_letter(52), "AZ")
  expect_equal(index_to_col_letter(53), "BA")
  expect_equal(index_to_col_letter(702), "ZZ")
})

test_that("col_letter_to_index and index_to_col_letter are inverses", {
  for (i in c(1, 5, 26, 27, 52, 100, 256, 702)) {
    expect_equal(col_letter_to_index(index_to_col_letter(i)), i)
  }
})

test_that("generate_col_names produces correct sequences", {
  expect_equal(generate_col_names(3), c("A", "B", "C"))
  expect_equal(generate_col_names(1), "A")
  cols26 <- generate_col_names(26)
  expect_equal(cols26[1], "A")
  expect_equal(cols26[26], "Z")
  cols28 <- generate_col_names(28)
  expect_equal(cols28[27], "AA")
  expect_equal(cols28[28], "AB")
})

test_that("generate_col_names rejects invalid input", {
  expect_error(generate_col_names(0))
  expect_error(generate_col_names(-1))
})

test_that("sanitize_sheet_name handles various inputs", {
  expect_equal(sanitize_sheet_name("Sheet 1"), "Sheet_1")
  expect_equal(sanitize_sheet_name("Non-Residents Tour Expenditure"),
               "NonResidents_Tour_Expenditure")
  expect_equal(sanitize_sheet_name("Final tables (detailed)"),
               "Final_tables_detailed")
  expect_equal(sanitize_sheet_name("Inter. Transport Services"),
               "Inter_Transport_Services")
  expect_equal(sanitize_sheet_name("HH Domestic Consumption"),
               "HH_Domestic_Consumption")
  # Already clean
  expect_equal(sanitize_sheet_name("Sheet1"), "Sheet1")
})

test_that("is_within_quotes detects quoted regions", {
  expect_false(is_within_quotes("A+B", 1))
  expect_true(is_within_quotes('A+"hello"+B', 5))
  expect_false(is_within_quotes('A+"hello"+B', 1))
  expect_false(is_within_quotes('A+"hello"+B', 10))
  # No quotes at all
  expect_false(is_within_quotes("ABC", 2))
})

test_that("parse_cell_address works for various formats", {
  p <- parse_cell_address("D10")
  expect_equal(p$col, "D")
  expect_equal(p$row, 10)

  p2 <- parse_cell_address("$AB$5")
  expect_equal(p2$col, "AB")
  expect_equal(p2$row, 5)

  p3 <- parse_cell_address("A1")
  expect_equal(p3$col, "A")
  expect_equal(p3$row, 1)

  # Column only (no row)
  p4 <- parse_cell_address("A")
  expect_equal(p4$col, "A")
  expect_true(is.na(p4$row))
})

test_that("expand_range_to_cells works for simple ranges", {
  cells <- expand_range_to_cells("A1:A3")
  expect_equal(cells, c("A1", "A2", "A3"))

  cells2 <- expand_range_to_cells("A1:B2")
  expect_equal(sort(cells2), sort(c("A1", "A2", "B1", "B2")))
})
