test_that("testsetup", {

  options("openxlsx2.datetimeCreated" = as.POSIXct("2023-07-20 23:32:14"))

  if (Sys.getenv("openxlsx2_testthat_fullrun") == "") {
    skip_on_ci()
  }

  skip_if_offline()
  expect_true(download_testfiles())
})
