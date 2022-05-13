test_that("Workbook class", {
  expect_null(assert_workbook(wb_workbook()))
})


test_that("wb_set_col_widths", {
# TODO use wb$wb_set_col_widths()

  wb <- wbWorkbook$new()
  wb <- wb$add_worksheet("test")
  wb$add_data("test", mtcars)

  # set column width to 12
  expect_silent(wb$set_col_widths("test", widths = 12L, cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" width=\"12\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # wrong sheet
  expect_error(wb$set_col_widths("test2", widths = 12L, cols = seq_along(mtcars)))

  # reset the column with, we do not provide an option ot remove the column entry
  expect_silent(wb$set_col_widths("test", cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" width=\"8.43\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # create column width for column 25
  expect_silent(wb$set_col_widths("test", cols = "Y", widths = 22))
  expect_equal(
    c("<col min=\"1\" max=\"11\" width=\"8.43\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
      "<col min=\"12\" max=\"24\" width=\"8.43\"/>",
      "<col min=\"25\" max=\"25\" width=\"22\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>"),
    wb$worksheets[[1]]$cols_attr
  )

  # a few more errors
  expect_error(wb$set_col_widths("test", cols = "Y", width = 1:2))
  expect_error(wb$set_col_widths("test", cols = "Y", hidden = 1:2))
})
