

test_that("Worksheet names", {

  ### test for names without special character
  sheetname <- "test"
  wb <- wb_add_worksheet(wb_workbook(), sheetname)
  expect_identical(sheetname, names(wb))

  # ### test for names with &
  # sheetname <- "S&P 500"
  # wb <- wb_add_worksheet(wb_workbook(), sheetname)
  # expect_identical(sheetname, names(wb))
  # expect_identical("S&amp;P 500", wb$sheet_names)
  #
  # ### test for names with <
  # sheetname <- "<24 h"
  # wb <- wb_add_worksheet(wb_workbook(), sheetname)
  #
  # expect_identical(sheetname, names(wb))
  # expect_identical("&lt;24 h",wb$sheet_names)
  #
  # ### test for names with >
  # sheetname <- ">24 h"
  # wb <- wb_add_worksheet(wb_workbook(), sheetname)
  # expect_identical(sheetname, names(wb))
  # expect_identical("&gt;24 h", wb$sheet_names)
  #
  # ### test for names with "
  # sheetname <- 'test "A"'
  # wb <- wb_add_worksheet(wb_workbook(), sheetname)
  # expect_identical(sheetname, names(wb))
  # expect_identical("test &quot;A&quot;", wb$sheet_names)
})

test_that("worksheet names with illegal characters not allowed", {
  pat <- "[Ii]llegal characters"
  wb <- wb_workbook()
  expect_error(wb$add_worksheet("S&P 500"),  pat)
  expect_error(wb$add_worksheet("<24 h"),    pat)
  expect_error(wb$add_worksheet(">24 h"),    pat)
  expect_error(wb$add_worksheet('test "A"'), pat)
})
