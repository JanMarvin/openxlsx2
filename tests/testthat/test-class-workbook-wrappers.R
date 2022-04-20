
# wb_remove_row_heights() -----------------------------------------------------

test_that("wb_remove_row_heights() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  wb$addWorksheet("a")
  wb$setRowHeights("a", 1:3, 20)
  params <- list(sheet = "sheet", rows = 2)
  expect_wrapper("removeRowHeights", "wb_remove_row_heights", wb = wb, params = params)
})
