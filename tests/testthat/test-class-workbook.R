test_that("Workbook class", {
  expect_null(assert_workbook(wb_workbook()))
})


test_that("setColWidths", {
# TODO use wb$setColWidths()

  wb <- wbWorkbook$new()
  wb$addWorksheet("test")
  writeData(wb, "test", mtcars)

  # set column width to 12
  expect_silent(setColWidths(wb, "test", widths = 12L, cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" width=\"12\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # wrong sheet
  expect_error(setColWidths(wb, "test2", widths = 12L, cols = seq_along(mtcars)))

  # reset the column with, we do not provide an option ot remove the column entry
  expect_silent(setColWidths(wb, "test", cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" width=\"8.43\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # create column width for column 25
  expect_silent(setColWidths(wb, "test", cols = "Y", widths = 22))
  expect_equal(
    c("<col min=\"1\" max=\"11\" width=\"8.43\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
      "<col min=\"12\" max=\"24\" width=\"8.43\"/>",
      "<col min=\"25\" max=\"25\" width=\"22\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>"),
    wb$worksheets[[1]]$cols_attr
  )

  # a few more errors
  expect_error(setColWidths(wb, "test", cols = "Y", width = 1:2))
  expect_error(setColWidths(wb, "test", cols = "Y", hidden = 1:2))
})
