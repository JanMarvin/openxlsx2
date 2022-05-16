
test_that("Worksheet names", {
  ### test for names without special character
  sheetname <- "test"
  wb <- wb_add_worksheet(wb_workbook(), sheetname)
  expect_identical(sheetname, names(wb))
})

test_that("worksheet names with illegal characters not allowed", {
  pat <- "[Ii]llegal characters"
  wb <- wb_workbook()
  expect_error(wb$add_worksheet("S&P 500"),  pat)
  expect_error(wb$add_worksheet("<24 h"),    pat)
  expect_error(wb$add_worksheet(">24 h"),    pat)
  expect_error(wb$add_worksheet('test "A"'), pat)
})
