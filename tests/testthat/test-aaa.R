test_that("testsetup", {
  skip_if_offline()
  expect_true(download_testfiles())
})
