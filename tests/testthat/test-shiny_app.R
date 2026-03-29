# =============================================================================
# Tests for the Shiny app using shinytest2
# =============================================================================

# Helper to find the project root
find_project_root <- function() {
  candidates <- c(
    ".",
    "../..",
    file.path(getwd(), "../..")
  )
  for (d in candidates) {
    if (file.exists(file.path(d, "app.R"))) return(normalizePath(d))
  }
  NULL
}

find_demo_file <- function() {
  candidates <- c(
    "inst/demo/sales_report_demo.xlsx",
    "../../inst/demo/sales_report_demo.xlsx"
  )
  for (f in candidates) {
    if (file.exists(f)) return(normalizePath(f))
  }
  NULL
}

test_that("App launches without errors", {
  skip_if_not_installed("shinytest2")
  skip_on_ci()

  app_dir <- find_project_root()
  skip_if(is.null(app_dir), "app.R not found")

  app <- shinytest2::AppDriver$new(
    app_dir = app_dir,
    name = "excel2r-basic",
    timeout = 15000,
    height = 800,
    width = 1200
  )
  on.exit(app$stop(), add = TRUE)

  # App should start without error
  expect_no_error(app$get_value(output = "upload_status"))
})

test_that("App processes uploaded demo file", {
  skip_if_not_installed("shinytest2")
  skip_on_ci()

  app_dir <- find_project_root()
  skip_if(is.null(app_dir), "app.R not found")

  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo Excel file not found")

  app <- shinytest2::AppDriver$new(
    app_dir = app_dir,
    name = "excel2r-upload",
    timeout = 30000,
    height = 800,
    width = 1200
  )
  on.exit(app$stop(), add = TRUE)

  # Upload file
  app$upload_file(upload = demo_file)

  # Wait for processing
  Sys.sleep(10)

  # Check that results are populated
  n_formulas <- app$get_value(output = "n_formulas")
  expect_true(!is.null(n_formulas))
})
