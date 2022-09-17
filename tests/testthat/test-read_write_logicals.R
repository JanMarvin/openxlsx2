test_that("TRUE, FALSE, NA", {
  curr_wd <- getwd()
  fileName <- file.path(tempdir(), "T_F_NA.xlsx")

  x <- iris
  x$Species <- as.character(x$Species)
  x$all_t <- TRUE
  x$all_f <- FALSE
  x$tf <- sample(c(TRUE, FALSE), size = 150, replace = TRUE)
  x$t_na <- sample(c(TRUE, NA), size = 150, replace = TRUE)
  x$f_na <- sample(c(FALSE, NA), size = 150, replace = TRUE)
  x$tf_na <- sample(c(TRUE, FALSE, NA), size = 150, replace = TRUE)

  wb <- write_xlsx(x, file = fileName, colNames = TRUE)

  y <- read_xlsx(fileName, sheet = 1, na.strings = "#N/A")
  expect_equal(x, y, ignore_attr = TRUE)

  ## T becomes false TRUE and NA exist in a columns
  expect_equal(x$t_na, y$t_na)
  expect_equal(x$f_na, y$f_na)

  expect_equal(is.na(x$f_na), is.na(y$f_na))
  expect_equal(is.na(x$tf_na), is.na(y$tf_na))

  ## From Workbook
  y <- read_xlsx(wb, sheet = 1)
  expect_equal(x, y, ignore_attr = TRUE)


  ## T becomes false TRUE and NA exist in a columns
  expect_equal(x$t_na, y$t_na)
  expect_equal(x$f_na, y$f_na)

  expect_equal(is.na(x$f_na), is.na(y$f_na))
  expect_equal(is.na(x$tf_na), is.na(y$tf_na))

  expect_equal(object = getwd(), curr_wd)
  unlink(fileName, recursive = TRUE, force = TRUE)
})
