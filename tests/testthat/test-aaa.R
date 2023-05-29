test_that("testsetup", {

  if (Sys.getenv("openxlsx2_testthat_fullrun") == "") {
    skip_on_ci()
  }

  skip_if_offline()
  expect_true(download_testfiles())
})
