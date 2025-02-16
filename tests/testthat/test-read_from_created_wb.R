
test_that("Reading from new workbook", {
  curr_wd <- getwd()

  wb <- wb_workbook()
  for (i in 1:4) {
    wb$add_worksheet(sprintf("Sheet %s", i))
  }


  ## colNames = TRUE, rowNames = TRUE
  wb$add_data(sheet = 1, x = mtcars, col_names = TRUE, row_names = TRUE, start_row = 10, start_col = 5)
  x <- read_xlsx(wb, 1, col_names = TRUE, row_names = TRUE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)


  ## colNames = TRUE, rowNames = FALSE
  wb$add_data(sheet = 2, x = mtcars, col_names = TRUE, row_names = FALSE, start_row = 10, start_col = 5)
  x <- read_xlsx(wb, sheet = 2, col_names = TRUE, row_names = FALSE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)
  expect_equal(object = colnames(mtcars), expected = colnames(x), ignore_attr = TRUE)

  ## colNames = FALSE, rowNames = TRUE
  wb$add_data(sheet = 3, x = mtcars, col_names = FALSE, row_names = TRUE, start_row = 2, start_col = 2)
  x <- read_xlsx(wb, sheet = 3, col_names = FALSE, row_names = TRUE)
  expect_equal(object = mtcars, expected = x, ignore_attr = TRUE)
  expect_equal(object = rownames(x), expected = rownames(mtcars))


  ## colNames = FALSE, rowNames = FALSE
  wb$add_data(sheet = 4, x = mtcars, col_names = FALSE, row_names = FALSE, start_row = 12, start_col = 1)
  x <- read_xlsx(wb, sheet = 4, col_names = FALSE, row_names = FALSE)
  expect_equal(object = x, expected = mtcars, ignore_attr = TRUE)

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
  wb_save(wb, file = fileName, overwrite = TRUE)

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

  got1 <- wb_to_df(wb, dims = "A2:D5", col_names = FALSE)
  got2 <- wb_to_df(wb, rows = 2:5, cols = 1:4, col_names = FALSE)
  expect_equal(got1, got2)

  got3 <- wb_to_df(wb, rows = c(2:3, 5:6), cols = c(1, 3:5), col_names = FALSE)
  expect_equal(dim(got3), c(4L, 4L))
  expect_equal(colnames(got3), c("A", "C", "D", "E"))
  expect_equal(rownames(got3), c("2", "3", "5", "6"))

  got4 <- wb_to_df(wb, rows = 1:5, col_names = FALSE)
  expect_equal(dim(got4), c(5L, 1L))

  got5 <- wb_to_df(wb, cols = 1:5, col_names = FALSE)
  expect_equal(dim(got5), c(2, 5))

  got6 <- wb_to_df(wb, start_row = 4, cols = 1:4, col_names = FALSE)
  expect_true(all(is.na(got6)))
  expect_equal(rownames(got6), "4")

})

test_that("read startCol", {

  wb <- wb_workbook()$add_worksheet()$add_data(x = cars, start_col = "E")

  got <- wb_to_df(wb, start_col = 1, col_names = FALSE)
  expect_equal(names(got), LETTERS[1:6])

  got <- wb_to_df(wb, start_col = "F", col_names = FALSE)
  expect_equal(names(got), LETTERS[6])

})

test_that("reading with multiple sections in freezePane works", {
  temp <- temp_xlsx()
  wb <- wb_workbook()$add_worksheet()
  wb$worksheets[[1]]$freezePane <- "<pane xSplit=\"7320\" ySplit=\"1640\"/><selection pane=\"topRight\"/><selection pane=\"bottomLeft\"/><selection pane=\"bottomRight\" activeCell=\"C5\" sqref=\"C5\"/>"
  wb$save(temp)
  expect_silent(wb <- wb_load(temp))
})

test_that("skipEmptyCols keeps empty named columns", {

  ## initialize empty cells
  na_mat <- matrix(NA, nrow = 22, ncol = 7)

  ## create artificial data with empty column
  dat    <- iris[1:20, ]
  dat$Species <- NA_real_
  dat <- rbind(dat, NA)

  ## create the workbook
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = na_mat, col_names = FALSE)$
    add_data(x = dat)

  got <- wb_to_df(wb, skip_empty_cols = TRUE)
  expect_equal(dat, got, ignore_attr = TRUE)

})

test_that("reading with pre defined types works", {

  dat <- data.frame(
    numeric = seq(-0.1, 0.1, by = 0.05),
    integer = sample(1:5, 5, TRUE),
    date = Sys.Date() - 0:4,
    datetime = Sys.time() - 0:4,
    character = letters[1:5],
    stringsAsFactors = FALSE
  )

  wb <- wb_workbook()$add_worksheet()$add_data(x = dat)

  got <- wb_to_df(wb)
  expect_equal(got, dat, ignore_attr = TRUE)

  types <- c(numeric = 1, integer = 1, date = 2, datetime = 3, character = 0)
  got <- wb_to_df(wb, types = types)
  expect_equal(got, dat, ignore_attr = TRUE)

})

test_that("wb_load contains path", {

  tmp <- temp_xlsx()
  wb_workbook()$add_worksheet()$add_worksheet()$save(tmp)
  wb_load(tmp)$remove_worksheet()$save()

  wb <- wb_load(tmp)

  exp <- c(`Sheet 1` = "Sheet 1")
  got <- wb$get_sheet_names()
  expect_equal(exp, got)

})

test_that("column names are not missing with col_names = FALSE", {
  dat <- data.frame(
    numeric = 1:2,
    character = c("a", "b"),
    logical = c(TRUE, FALSE),
    stringsAsFactors = FALSE
  )

  wb <- wb_workbook()$add_worksheet()$add_data(x = dat)
  got <- as.character(wb_to_df(wb, col_names = FALSE)[2, ])
  exp <- c("1", "a", "TRUE")
  expect_equal(exp, got)

})

test_that("check_names works", {

  dd <- data.frame(
    "a and b"  = 1:2,
    "a-and-b" = 3:4,
    check.names = FALSE
  )

  wb <- write_xlsx(x = dd)

  exp <- c("a and b", "a-and-b")
  got <- names(wb_to_df(wb, check_names = FALSE))
  expect_equal(exp, got)

  exp <- c("a.and.b", "a.and.b.1")
  got <- names(wb_to_df(wb, check_names = TRUE))
  expect_equal(exp, got)

})

test_that("shared formulas are handled", {

  mm <- matrix(1:5, ncol = 5, nrow = 5)
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mm)$
    add_formula(x = "SUM($A2:A2) + B$1", dims = "A8:E12", shared = TRUE)

  exp <- structure(
    c("SUM($A4:A4) + B$1", "SUM($A5:A5) + B$1", "SUM($A6:A6) + B$1",
      "SUM($A4:B4) + C$1", "SUM($A5:B5) + C$1", "SUM($A6:B6) + C$1",
      "SUM($A4:C4) + D$1", "SUM($A5:C5) + D$1", "SUM($A6:C6) + D$1"),
    dim = c(3L, 3L),
    dimnames = list(c("10", "11", "12"), c("A", "B", "C"))
  )
  got <- as.matrix(wb_to_df(wb, dims = "A10:C12", col_names = FALSE, show_formula = TRUE))
  expect_equal(exp, got)

  # shared formula w/o A1 notation
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mm)$
    add_formula(x = "TODAY()", dims = "A8:E12", shared = TRUE)

  exp <- rep("TODAY()", 9)
  got <- unname(unlist(wb_to_df(wb, dims = "A10:C12", col_names = FALSE, show_formula = TRUE)))
  expect_equal(exp, got)

  # a bunch of mixed shared formulas
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mm)$
    add_formula(x = "SUM($A2:A2) + B$1", dims = "A8:E9", shared = TRUE)$
    add_formula(x = "SUM($A2:A2)", dims = "A10:E11", shared = TRUE)$
    add_formula(x = "A2", dims = "A12:E13", shared = TRUE)

  exp <- c("SUM($A2:B2) + C$1", "SUM($A3:B3) + C$1", "SUM($A2:B2)", "SUM($A3:B3)",
           "B2", "B3", "SUM($A2:C2) + D$1", "SUM($A3:C3) + D$1", "SUM($A2:C2)",
           "SUM($A3:C3)", "C2", "C3", "SUM($A2:D2) + E$1", "SUM($A3:D3) + E$1",
           "SUM($A2:D2)", "SUM($A3:D3)", "D2", "D3")
  got <- unname(unlist(wb_to_df(wb, dims = "B8:D13", col_names = FALSE, show_formula = TRUE)))
  expect_equal(exp, got)

  # make sure that replacements work as expected
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mm)$
    add_formula(x = "A2 + B2", dims = "A12:E13", shared = TRUE)

  exp <- c("A2 + B2", "A3 + B3", "B2 + C2", "B3 + C3", "C2 + D2", "C3 + D3", "D2 + E2", "D3 + E3")
  got <- unname(unlist(wb_to_df(wb, dims = "A12:D13", col_names = FALSE, show_formula = TRUE)))
  expect_equal(exp, got)

})

test_that("reading with blank rows works (#1272)", {

  wb <- wb_workbook()$add_worksheet()$
    add_data(dims = "A1", x = 1)$
    add_data(dims = "A4:E4", x = rep(1, 5))

  rr <- wb$worksheets[[1]]$sheet_data$row_attr

  rr <- create_char_dataframe(names(rr), n = 2)
  rr$r <- as.character(2:3)

  wb$worksheets[[1]]$sheet_data$row_attr <- rbind(
    wb$worksheets[[1]]$sheet_data$row_attr,
    rr,
    stringAsFactors = FALSE
  )

  got <- wb$to_df(dims = "A2:E3", col_names = FALSE)

  expect_equal(c(2L, 5L), dim(got))

})
