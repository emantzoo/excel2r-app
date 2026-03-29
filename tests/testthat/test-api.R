# =============================================================================
# Tests for public API functions
# =============================================================================

test_that("migrate() produces output files in CSV mode", {
  demo <- find_demo_file()
  skip_if(is.null(demo), "Demo workbook not found")

  tmp <- file.path(tempdir(), "test_migrate_csv")
  on.exit(unlink(tmp, recursive = TRUE))

  result <- migrate(demo, tmp, mode = "csv", verify = FALSE)

  expect_true(file.exists(file.path(tmp, "generated_script.R")))
  expect_true(dir.exists(file.path(tmp, "data")))
  csv_files <- list.files(file.path(tmp, "data"), pattern = "\\.csv$")
  expect_true(length(csv_files) > 0)
  expect_true(!is.null(result$report))
})

test_that("migrate() produces script in Excel mode", {
  demo <- find_demo_file()
  skip_if(is.null(demo), "Demo workbook not found")

  tmp <- file.path(tempdir(), "test_migrate_excel")
  on.exit(unlink(tmp, recursive = TRUE))

  result <- migrate(demo, tmp, mode = "excel", verify = FALSE)

  expect_true(file.exists(file.path(tmp, "generated_script.R")))
  expect_false(dir.exists(file.path(tmp, "data")))
})

test_that("process() returns expected structure", {
  demo <- find_demo_file()
  skip_if(is.null(demo), "Demo workbook not found")

  result <- process(demo)

  expect_true(is.character(result$script))
  expect_true(is.data.frame(result$report))
  expect_true(nrow(result$report) > 0)
  expect_true("R_Code" %in% colnames(result$report))
})

test_that("verify() returns summary and mismatches", {
  demo <- find_demo_file()
  skip_if(is.null(demo), "Demo workbook not found")

  result <- process(demo, mode = "excel")
  v <- verify(demo, result)

  expect_true(is.list(v$summary))
  expect_true(v$summary$total > 0)
  expect_true(is.data.frame(v$mismatches))
})

test_that("supported_functions() returns 50+ functions", {
  fns <- supported_functions()
  expect_true(length(fns) >= 50)
  expect_true("SUM" %in% fns)
  expect_true("VLOOKUP" %in% fns)
})
