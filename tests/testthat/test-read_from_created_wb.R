
context("Reading from wb object is identical to reading from file")

test_that("Reading from new workbook", {
  curr_wd <- getwd()

  wb <- createWorkbook()
  for (i in 1:4) {
    addWorksheet(wb, sprintf("Sheet %s", i))
  }


  ## colNames = TRUE, rowNames = TRUE
  writeData(wb, sheet = 1, x = mtcars, colNames = TRUE, rowNames = TRUE, startRow = 10, startCol = 5)
  x <- read.xlsx(wb, 1, colNames = TRUE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)


  ## colNames = TRUE, rowNames = FALSE
  writeData(wb, sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5)
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  expect_equal(object = colnames(mtcars), expected = colnames(x), check.attributes = FALSE)

  ## colNames = FALSE, rowNames = TRUE
  writeData(wb, sheet = 3, x = mtcars, colNames = FALSE, rowNames = TRUE, startRow = 2, startCol = 2)
  x <- read.xlsx(wb, sheet = 3, colNames = FALSE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  expect_equal(object = rownames(mtcars), expected = rownames(x))


  ## colNames = FALSE, rowNames = FALSE
  writeData(wb, sheet = 4, x = mtcars, colNames = FALSE, rowNames = FALSE, startRow = 12, startCol = 1)
  x <- read.xlsx(wb, sheet = 4, colNames = FALSE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)

  expect_equal(object = getwd(), curr_wd)
  rm(wb)
})

test_that("Reading NAs and NaN values", {
  fileName <- file.path(tempdir(), "NaN.xlsx")
  na.string <- "*"

  ## data
  a <- data.frame(
    X  = c(-pi / 0, NA, NaN),
    Y  = letters[1:3],
    Z  = c(pi / 0, 99, NaN),
    Z2 = c(1, NaN, NaN),
    stringsAsFactors = FALSE
  )

  b <- a
  b[b == -Inf] <- NaN
  b[b == Inf] <- NaN

  c <- b
  is_na <- sapply(c, is.na)
  is_nan <- sapply(c, is.nan)
  c[is_na & !is_nan] <- na.string
  is_nan_after <- sapply(c, is.nan)
  c[is_nan & !is_nan_after] <- NA

  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  writeData(wb, 1, a, keepNA = FALSE)

  addWorksheet(wb, "Sheet 2")
  writeData(wb, 2, a, keepNA = TRUE)

  addWorksheet(wb, "Sheet 3")
  writeData(wb, 3, a, keepNA = TRUE, na.string = na.string)

  saveWorkbook(wb, file = fileName, overwrite = TRUE)

  ## from file
  expected_df <- structure(list(
    X = c(NA_real_, NA_real_, NA_real_),
    Y = c("a", "b", "c"),
    Z = c(NA, 99, NA),
    Z2 = c(1, NA, NA)
  ),
    .Names = c("X", "Y", "Z", "Z2"),
    row.names = c(NA, 3L), class = "data.frame"
  )

  expect_equal(expect_warning(read.xlsx(fileName)), a, check.attributes=FALSE)
  unlink(fileName, recursive = TRUE, force = TRUE)
})
