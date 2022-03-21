

test_that("clone Worksheet with data", {
  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  writeData(wb, "Sheet 1", 1)
  wb$cloneWorksheet("Sheet 2", "Sheet 1")

  file_name <- system.file("extdata", "cloneWorksheetExample.xlsx", package = "openxlsx2")
  refwb <- loadWorkbook(file = file_name)

  expect_equal(names(wb), names(refwb))
  expect_equal(worksheetOrder(wb), worksheetOrder(refwb))
})

test_that("clone empty Worksheet", {
  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$cloneWorksheet("Sheet 2", "Sheet 1")

  file_name <- system.file("extdata", "cloneEmptyWorksheetExample.xlsx", package = "openxlsx2")
  refwb <- loadWorkbook(file = file_name)

  expect_equal(names(wb), names(refwb))
  expect_equal(worksheetOrder(wb), worksheetOrder(refwb))
})
