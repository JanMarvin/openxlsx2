test_that("clone Worksheet with data", {
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_data("Sheet 1", 1)
  wb$clone_worksheet("Sheet 1", "Sheet 2")

  file_name <- system.file("extdata", "cloneWorksheetExample.xlsx", package = "openxlsx2")
  refwb <- wb_load(file = file_name)

  expect_equal(names(wb), names(refwb))
  expect_equal(wb_get_order(wb), wb_get_order(refwb))
})


test_that("clone empty Worksheet", {
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$clone_worksheet("Sheet 1", "Sheet 2")

  file_name <- system.file("extdata", "cloneEmptyWorksheetExample.xlsx", package = "openxlsx2")
  refwb <- wb_load(file = file_name)

  expect_equal(names(wb), names(refwb))
  expect_equal(wb_get_order(wb), wb_get_order(refwb))
})
