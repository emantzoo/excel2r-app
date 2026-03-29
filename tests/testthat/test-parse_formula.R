# =============================================================================
# Tests for R/parse_formula.R
# =============================================================================

test_that("find_matching_paren handles simple cases", {
  expect_equal(find_matching_paren("SUM(A1)", 4), 7)
  expect_equal(find_matching_paren("SUM(A1:A10)", 4), 11)
  expect_equal(find_matching_paren("()", 1), 2)
})

test_that("find_matching_paren handles nested parentheses", {
  # "SUM(IF(A1,B1,C1))" — outer ) at pos 17, inner ) at pos 16
  expect_equal(find_matching_paren("SUM(IF(A1,B1,C1))", 4), 17)
  expect_equal(find_matching_paren("SUM(IF(A1,B1,C1))", 7), 16)
  expect_equal(find_matching_paren("A*(B+(C*D))", 3), 11)
})

test_that("find_matching_paren handles strings with parens", {
  # The paren inside quotes should not count
  expect_equal(find_matching_paren('SUM("(test)",A1)', 4), 16)
})

test_that("find_matching_paren returns NA for unmatched", {
  expect_true(is.na(find_matching_paren("SUM(A1", 4)))
})

test_that("split_function_args splits at top-level commas", {
  args <- split_function_args("A1, B1, C1")
  expect_equal(args, c("A1", "B1", "C1"))
})

test_that("split_function_args respects nested parens", {
  args <- split_function_args("IF(A1>0,B1,0), C1")
  expect_equal(length(args), 2)
  expect_equal(args[1], "IF(A1>0,B1,0)")
  expect_equal(args[2], "C1")
})

test_that("split_function_args respects quoted strings with commas", {
  args <- split_function_args('A1:A10, "hello,world", B1')
  expect_equal(length(args), 3)
  expect_equal(args[2], '"hello,world"')
})

test_that("split_function_args handles single argument", {
  args <- split_function_args("A1:A10")
  expect_equal(args, "A1:A10")
})

test_that("split_function_args handles empty string", {
  args <- split_function_args("")
  expect_equal(length(args), 0)
})

test_that("find_function_calls finds simple functions", {
  calls <- find_function_calls("SUM(A1:A10)")
  expect_equal(length(calls), 1)
  expect_equal(calls[[1]]$name, "SUM")
  expect_equal(calls[[1]]$args_str, "A1:A10")
})

test_that("find_function_calls finds nested functions", {
  calls <- find_function_calls("SUM(IF(A1>0,B1,0),C1)")
  # Should find both SUM and IF
  func_names <- sapply(calls, function(x) x$name)
  expect_true("SUM" %in% func_names)
  expect_true("IF" %in% func_names)
})

test_that("find_function_calls finds multiple independent functions", {
  calls <- find_function_calls("SUM(A1)+MAX(B1)")
  func_names <- sapply(calls, function(x) x$name)
  expect_true("SUM" %in% func_names)
  expect_true("MAX" %in% func_names)
})

test_that("find_function_calls returns empty for no functions", {
  calls <- find_function_calls("A1+B1*C1")
  expect_equal(length(calls), 0)
})

test_that("extract_ranges finds simple ranges", {
  result <- extract_ranges("SUM(D11:D12)")
  expect_equal(length(result$placeholder_map), 1)
  expect_true(grepl("<RANGE_1>", result$placeholder_formula))
  expect_equal(unname(result$placeholder_map[[1]]), "D11:D12")
})

test_that("extract_ranges finds multiple ranges", {
  result <- extract_ranges("SUM(A1:A10)+SUM(B1:B5)")
  expect_equal(length(result$placeholder_map), 2)
})

test_that("extract_ranges handles cross-sheet ranges", {
  result <- extract_ranges("'Sheet 1'!A1:A10")
  expect_equal(length(result$placeholder_map), 1)
})

test_that("extract_ranges handles whole-column ranges", {
  result <- extract_ranges("SUMIFS(E:E,B:B,\"YES\")")
  expect_equal(length(result$placeholder_map), 2)
})

test_that("extract_ranges ignores colons inside quotes", {
  result <- extract_ranges('IF(A1="10:30",B1,C1)')
  # Should not detect "10:30" as a range
  expect_equal(length(result$placeholder_map), 0)
})
