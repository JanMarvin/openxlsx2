
test_that("Reading from new workbook", {
  curr_wd <- getwd()

  wb <- wb_workbook()
  for (i in 1:4) {
    wb$add_worksheet(sprintf("Sheet %s", i))
  }


  ## colNames = TRUE, rowNames = TRUE
  wb$add_data(sheet = 1, x = mtcars, colNames = TRUE, rowNames = TRUE, startRow = 10, startCol = 5)
  x <- read_xlsx(wb, 1, colNames = TRUE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)


  ## colNames = TRUE, rowNames = FALSE
  wb$add_data(sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5)
  x <- read_xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)
  expect_equal(object = colnames(mtcars), expected = colnames(x), ignore_attr = TRUE)

  ## colNames = FALSE, rowNames = TRUE
  wb$add_data(sheet = 3, x = mtcars, colNames = FALSE, rowNames = TRUE, startRow = 2, startCol = 2)
  x <- read_xlsx(wb, sheet = 3, colNames = FALSE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)
  expect_equal(object = rownames(mtcars), expected = rownames(x))


  ## colNames = FALSE, rowNames = FALSE
  wb$add_data(sheet = 4, x = mtcars, colNames = FALSE, rowNames = FALSE, startRow = 12, startCol = 1)
  x <- read_xlsx(wb, sheet = 4, colNames = FALSE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)

  expect_equal(object = getwd(), curr_wd)
  rm(wb)
})

test_that("Reading NAs and NaN values", {
  fileName <- file.path(tempdir(), "NaN.xlsx")
  na.string <- "*"

  wb <- wb_workbook()

  ## data
  A <- data.frame(
    X  = c(-pi / 0, NA, NaN),
    Y  = letters[1:3],
    Z  = c(pi / 0, 99, NaN),
    Z2 = c(1, NaN, NaN),
    stringsAsFactors = FALSE
  )

  wb$add_worksheet("Sheet 1")$add_data(1, A)

  # not used
  B <- A
  B[B == -Inf] <- NaN
  B[B == Inf] <- NaN

  wb$add_worksheet("Sheet 2")$add_data(2, B)

  # not used
  C <- B
  is_na <- sapply(C, is.na)
  is_nan <- sapply(C, is.nan)
  C[is_na & !is_nan] <- na.string
  is_nan_after <- sapply(C, is.nan)
  C[is_nan & !is_nan_after] <- NA

  wb$add_worksheet("Sheet 3")$add_data(3, C)

  # write file
  wb_save(wb, path = fileName, overwrite = TRUE)

  exp <- data.frame(
    X  = c("#NUM!", NA, "#VALUE!"),
    Y  = letters[1:3],
    Z  = c("#NUM!", 99, "#VALUE!"),
    Z2 = c(1, "#VALUE!", "#VALUE!"),
    stringsAsFactors = FALSE
  )

  expect_equal(read_xlsx(fileName), exp, ignore_attr = TRUE)
  unlink(fileName, recursive = TRUE, force = TRUE)
})

test_that("dims != rows & cols", {

  wb <- wb_workbook()$
    add_worksheet("sheet1")$
    add_data(1, data.frame("A" = 1))

  got1 <- wb_to_df(wb, dims = "A2:D5", colNames = FALSE)
  got2 <- wb_to_df(wb, rows = 2:5, cols = 1:4, colNames = FALSE)
  expect_equal(got1, got2)

  got3 <- wb_to_df(wb, rows = c(2:3, 5:6), cols = c(1, 3:5), colNames = FALSE)
  expect_equal(c(4,4), dim(got3))
  expect_equal(c("A", "C", "D", "E"), colnames(got3))
  expect_equal(c("2", "3", "5", "6"), rownames(got3))

  got4 <- wb_to_df(wb, rows = 1:5, colNames = FALSE)
  expect_equal(c(5, 1), dim(got4))

  got5 <- wb_to_df(wb, cols = 1:5, colNames = FALSE)
  expect_equal(c(2, 5), dim(got5))

  got6 <- wb_to_df(wb, startRow = 4, cols = 1:4, colNames = FALSE)
  expect_true(all(is.na(got6)))
  expect_equal("4", rownames(got6))

})

test_that("read startCol", {

  wb <- wb_workbook()$add_worksheet()$add_data(x = cars, startCol = "E")

  got <- wb_to_df(wb, startCol = 1, colNames = FALSE)
  expect_equal(LETTERS[1:6], names(got))

  got <- wb_to_df(wb, startCol = "F", colNames = FALSE)
  expect_equal(LETTERS[6], names(got))

})
