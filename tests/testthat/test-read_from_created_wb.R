
test_that("Reading from new workbook", {
  curr_wd <- getwd()

  wb <- wb_workbook()
  for (i in 1:4) {
    wb$addWorksheet(sprintf("Sheet %s", i))
  }


  ## colNames = TRUE, rowNames = TRUE
  writeData(wb, sheet = 1, x = mtcars, colNames = TRUE, rowNames = TRUE, startRow = 10, startCol = 5)
  x <- read.xlsx(wb, 1, colNames = TRUE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)


  ## colNames = TRUE, rowNames = FALSE
  writeData(wb, sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5)
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)
  expect_equal(object = colnames(mtcars), expected = colnames(x), ignore_attr = TRUE)

  ## colNames = FALSE, rowNames = TRUE
  writeData(wb, sheet = 3, x = mtcars, colNames = FALSE, rowNames = TRUE, startRow = 2, startCol = 2)
  x <- read.xlsx(wb, sheet = 3, colNames = FALSE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)
  expect_equal(object = rownames(mtcars), expected = rownames(x))


  ## colNames = FALSE, rowNames = FALSE
  writeData(wb, sheet = 4, x = mtcars, colNames = FALSE, rowNames = FALSE, startRow = 12, startCol = 1)
  x <- read.xlsx(wb, sheet = 4, colNames = FALSE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)

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

  wb <- wb_workbook()

  wb$addWorksheet("Sheet 1")
  writeData(wb, 1, a, keepNA = FALSE)

  wb$addWorksheet("Sheet 2")
  writeData(wb, 2, a, keepNA = TRUE)

  wb$addWorksheet("Sheet 3")
  writeData(wb, 3, a, keepNA = TRUE, na.string = na.string)

  wb_save(wb, path = fileName, overwrite = TRUE)

  expect_equal(read.xlsx(fileName), a, ignore_attr = TRUE)
  unlink(fileName, recursive = TRUE, force = TRUE)
})
