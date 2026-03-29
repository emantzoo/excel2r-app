# Run all tests
# Usage: Rscript run_tests.R

# Source all modules
for (f in list.files("R", pattern = "\\.R$", full.names = TRUE)) {
  source(f)
}

# Run testthat
library(testthat)
test_dir("tests/testthat", reporter = "summary")
