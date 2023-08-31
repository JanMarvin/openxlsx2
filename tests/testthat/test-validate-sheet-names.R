
test_that("Worksheet names", {
  ### test for names without special character
  sheetname <- "test"
  wb <- wb_add_worksheet(wb_workbook(), sheetname)
  expect_identical(sheetname, names(wb$get_sheet_names()))
})

test_that("legal characters are allowed", {
  wb <- wb_workbook()
  expect_error(wb$add_worksheet("S&P 500"),  NA)
  expect_error(wb$add_worksheet("<24 h"),    NA)
  expect_error(wb$add_worksheet(">24 h"),    NA)
  expect_error(wb$add_worksheet('test "A"'), NA)
})

test_that("illegal characters are not allowed", {
  wb <- wb_workbook()
  expect_warning(wb$add_worksheet("forward\\slash"), "illegal")
  expect_warning(wb$add_worksheet("back/slash"),     "illegal")
  expect_warning(wb$add_worksheet("question?mark"),  "illegal")
  expect_warning(wb$add_worksheet("asterik*"),       "illegal")
  expect_warning(wb$add_worksheet("colon:"),         "illegal")
  expect_warning(wb$add_worksheet("open[bracket"),   "illegal")
  expect_warning(wb$add_worksheet("closed]bracket"), "illegal")
})

test_that("validate sheet", {
  wb <- wb_workbook()
  expect_error(wb$add_worksheet(1:2), "length 1")
  expect_error(wb$add_worksheet(NA), "NA")
  expect_error(wb$add_worksheet(1.5), "integer")
  expect_error(wb$add_worksheet(-1), "positive")
  expect_error(wb$add_worksheet(1), NA)
  expect_error(wb$add_worksheet(1), "index")
  expect_warning(wb$add_worksheet(""), "at least 1 character")
  expect_error(wb$add_worksheet("0123456789012345789012345678901"), NA)
  expect_error(expect_warning(wb$add_worksheet("01234567890123457890123456789012"), "31"), "Cannot shorten")
  expect_error(wb$add_worksheet("my_sheet"), NA)
  expect_warning(wb$add_worksheet("my_sheet"), "already exists")
  expect_error(wb$add_worksheet(" "), NA)
})
