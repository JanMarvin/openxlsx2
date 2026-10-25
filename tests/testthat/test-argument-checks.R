test_that("wbWorkbook methods check their arguments", {

  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$set_sheetview(zoom_scale = 1:3), "`zoom_scale` must be a single value")
  expect_error(wb$set_page_setup(copies = 1:3), "`copies` must be a single value")
  expect_error(wb$set_bookview(window_height = c(1, 2)), "`window_height` must be a single value")
  expect_error(wb$add_worksheet(zoom = c(100, 120)), "`zoom` must be a single value")

  expect_error(wb_add_data(wb, x = 1, start_col = 1.5), "`start_col` must be a single whole number")
  expect_error(wb_add_data(wb, x = 1, start_row = c(1, 2)), "`start_row` must be a single whole number")
  expect_error(wb_add_data(wb, x = 1, col_names = NA), "`col_names` must be TRUE or FALSE")
  expect_error(wb_add_data(wb, x = 1, col_names = c(TRUE, FALSE)), "`col_names` must be TRUE or FALSE")
  expect_error(wb_add_data_table(wb, x = data.frame(a = 1), with_filter = "yes"), "`with_filter` must be TRUE or FALSE")
  expect_error(wb_add_formula(wb, x = "1+1", array = NA), "`array` must be TRUE or FALSE")
  expect_error(wb_add_fill(wb, dims = 1), "`dims` must be a character vector")
  expect_error(wb_add_font(wb, dims = NA_character_), "`dims` must be a character vector")
  expect_error(wb_add_data(wb, x = 1, sheet = mean), "sheet index")

  expect_error(wb_add_formula(wb, x = "1+1", array = 2), "`array` must be TRUE or FALSE")

  expect_silent(wb_add_data(wb, x = 1, start_col = "B"))
  expect_silent(wb_add_formula(wb, x = "1+1", array = 1, cm = 0L))
  expect_silent(wb$set_sheetview(zoom_scale = 120))
  expect_silent(wb$set_col_widths(cols = 1:2, widths = c(10, 20), hidden = c(TRUE, FALSE)))

})
