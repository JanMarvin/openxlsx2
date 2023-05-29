test_that("testsetup", {
  skip_on_ci()
  skip_if_offline()
  expect_true(download_testfiles())
})
