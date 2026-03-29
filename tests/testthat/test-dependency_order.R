# =============================================================================
# Tests for R/dependency_order.R
# =============================================================================

# --- Kahn's sort ---

test_that("kahns_sort handles simple linear chain", {
  nodes <- c("A", "B", "C")
  deps <- list(A = character(0), B = c("A"), C = c("B"))
  result <- kahns_sort(nodes, deps)
  expect_equal(result, c("A", "B", "C"))
})

test_that("kahns_sort handles no dependencies", {
  nodes <- c("A", "B", "C")
  deps <- list(A = character(0), B = character(0), C = character(0))
  result <- kahns_sort(nodes, deps)
  expect_equal(length(result), 3)
  expect_true(all(nodes %in% result))
})

test_that("kahns_sort handles diamond dependency", {
  # A -> B, A -> C, B -> D, C -> D
  nodes <- c("A", "B", "C", "D")
  deps <- list(A = character(0), B = c("A"), C = c("A"), D = c("B", "C"))
  result <- kahns_sort(nodes, deps)
  expect_equal(result[1], "A")
  expect_equal(result[4], "D")
})

test_that("kahns_sort detects cycles", {
  nodes <- c("A", "B", "C")
  deps <- list(A = c("C"), B = c("A"), C = c("B"))
  result <- kahns_sort(nodes, deps)
  expect_null(result)
})

# --- determine_execution_order ---

test_that("execution order handles single sheet, no deps", {
  data <- data.frame(
    Sheet = c("S1", "S1"),
    Cell = c("A1", "B1"),
    Formula = c("5", "10"),
    stringsAsFactors = FALSE
  )
  result <- determine_execution_order(data)
  expect_equal(result$sheet_order, "S1")
  expect_equal(length(result$cell_orders[["S1"]]), 2)
})

test_that("execution order respects cell dependencies", {
  data <- data.frame(
    Sheet = c("S1", "S1", "S1"),
    Cell = c("A1", "A2", "A3"),
    Formula = c("5", "A1+1", "A2+A1"),
    stringsAsFactors = FALSE
  )
  result <- determine_execution_order(data)
  cell_order <- result$cell_orders[["S1"]]

  # A1 must come before A2 and A3
  pos_a1 <- which(cell_order == "A1")
  pos_a2 <- which(cell_order == "A2")
  pos_a3 <- which(cell_order == "A3")
  expect_true(pos_a1 < pos_a2)
  expect_true(pos_a1 < pos_a3)
  # A2 must come before A3 (since A3 depends on A2)
  expect_true(pos_a2 < pos_a3)
})

test_that("execution order handles cross-sheet dependencies", {
  data <- data.frame(
    Sheet = c("Sheet1", "Sheet2"),
    Cell = c("A1", "A1"),
    Formula = c("5", "'Sheet1'!A1+10"),
    stringsAsFactors = FALSE
  )
  result <- determine_execution_order(data)

  # Sheet1 should come before Sheet2
  pos_s1 <- which(result$sheet_order == "Sheet1")
  pos_s2 <- which(result$sheet_order == "Sheet2")
  expect_true(pos_s1 < pos_s2)
})

test_that("execution order handles range dependencies", {
  data <- data.frame(
    Sheet = c("S1", "S1", "S1", "S1"),
    Cell = c("A1", "A2", "A3", "A4"),
    Formula = c("5", "10", "15", "SUM(A1:A3)"),
    stringsAsFactors = FALSE
  )
  result <- determine_execution_order(data)
  cell_order <- result$cell_orders[["S1"]]

  # A4 depends on A1, A2, A3 so should be last
  expect_equal(cell_order[length(cell_order)], "A4")
})

test_that("execution order warns on cycles instead of erroring", {
  data <- data.frame(
    Sheet = c("S1", "S1"),
    Cell = c("A1", "A2"),
    Formula = c("A2+1", "A1+1"),
    stringsAsFactors = FALSE
  )
  result <- determine_execution_order(data)
  # Should have a warning about the cycle
  expect_true(length(result$warnings) > 0)
  expect_true(any(grepl("Cycle", result$warnings)))
})

test_that("execution order handles empty data", {
  data <- data.frame(
    Sheet = character(0),
    Cell = character(0),
    Formula = character(0),
    stringsAsFactors = FALSE
  )
  result <- determine_execution_order(data)
  expect_equal(length(result$sheet_order), 0)
})
