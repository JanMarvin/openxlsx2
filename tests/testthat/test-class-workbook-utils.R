test_that("waivers works", {
  wb <- wb_workbook()
  expect_identical(wb$.__enclos_env__$private$current_sheet, 0L)
  expect_error(wb$add_worksheet(), NA)
  expect_identical(wb$.__enclos_env__$private$current_sheet, 1L)
  expect_error(wb$add_worksheet(), NA)
  expect_identical(wb$.__enclos_env__$private$current_sheet, 2L)
})
