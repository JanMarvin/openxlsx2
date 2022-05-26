
test_that("Worksheet names", {
  ### test for names without special character
  sheetname <- "test"
  wb <- wb_add_worksheet(wb_workbook(), sheetname)
  expect_identical(sheetname, names(wb))
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
  expect_error(wb$add_worksheet("forward\\slash"), "illegal")
  expect_error(wb$add_worksheet("back/slash"),     "illegal")
  expect_error(wb$add_worksheet("question?mark"),  "illegal")
  expect_error(wb$add_worksheet("asterik:"),       "illegal")
  expect_error(wb$add_worksheet("colon:"),         "illegal")
  expect_error(wb$add_worksheet("open[bracket"),   "illegal")
  expect_error(wb$add_worksheet("closed]bracket"), "illegal")
})
