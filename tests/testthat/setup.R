# Setup: source all R modules before tests run
r_dir <- file.path(dirname(dirname(getwd())), "R")
if (!dir.exists(r_dir)) {
  r_dir <- file.path(getwd(), "..", "..", "R")
}
if (!dir.exists(r_dir)) {
  r_dir <- "R"
}

for (f in list.files(r_dir, pattern = "\\.R$", full.names = TRUE)) {
  source(f)
}
