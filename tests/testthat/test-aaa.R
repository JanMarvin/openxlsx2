test_that("testsetup", {

  options("openxlsx2.datetimeCreated" = as.POSIXct("2023-07-20 23:32:14", tz = "UTC"))

  wb <- wb_workbook()

  exp <- "2023-07-20T23:32:14Z"
  got <- wb$get_properties()[["datetime_created"]]
  expect_equal(exp, got)

})
