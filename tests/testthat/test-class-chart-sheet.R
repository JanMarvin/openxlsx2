testsetup()

test_that("class wbChartSheet works", {

  wb <- wb_workbook()$add_chartsheet()
  expect_error(wb$add_data(x = mtcars), "Cannot write to chart sheet.")

})
