test_that("xl_open() works", {
  wb <- wb_workbook()$add_worksheet(1)
  # throw a warning
  expect_warning(expect_s3_class(wb$open(interactive = FALSE), c("R6", "wbWorkbook")))
  # should not change the path object
  expect_identical(wb$path, character())
})
