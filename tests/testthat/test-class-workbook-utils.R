testsetup()

test_that("waivers works with $add_worksheet()", {
  wb <- wb_workbook()
  expect_identical(wb$.__enclos_env__$private$current_sheet, 0L)
  expect_error(wb$add_worksheet(), NA)
  expect_identical(wb$.__enclos_env__$private$current_sheet, 1L)
  expect_error(wb$add_worksheet(), NA)
  expect_identical(wb$.__enclos_env__$private$current_sheet, 2L)
})

test_that("waivers work with $add_data()", {
  wb <- wb_workbook()
  expect_error(wb$add_worksheet()$add_data(x = data.frame(a = 1)), NA)
  expect_identical(wb$.__enclos_env__$private$current_sheet, 1L)
  expect_error(wb$add_worksheet()$add_data(x = data.frame(a = 1)), NA)
  expect_identical(wb$.__enclos_env__$private$current_sheet, 2L)
})
