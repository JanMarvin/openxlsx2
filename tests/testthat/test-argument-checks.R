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

test_that("cell style, ignore error and protection arguments are checked", {

  wb <- wb_workbook()$add_worksheet()$add_data(x = 1)

  expect_error(wb$add_cell_style(dims = "A1", wrap_text = c("1", "0")), "`wrap_text` must be a single value")
  expect_error(wb$add_ignore_error(dims = "A1", formula = "yes"), "`formula` must be TRUE or FALSE")
  expect_error(wb$protect_worksheet(protect = NA), "`protect` must be TRUE or FALSE")
  expect_error(wb$protect_worksheet(password = 1234), "`password` must be a single string")
  expect_error(wb$protect(lock_structure = "yes"), "`lock_structure` must be TRUE or FALSE")
  expect_error(wb$protect(password = c("a", "b")), "`password` must be a single string")

  expect_silent(wb$add_ignore_error(dims = "A1", number_stored_as_text = 1))
  expect_silent(wb$protect_worksheet(password = "pw"))

})

test_that("image, form control and page setup arguments are checked", {

  wb <- wb_workbook()$add_worksheet()
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  expect_error(wb$add_image(file = img, width = c(1, 2)), "`width` must be a single value")
  expect_error(wb$add_form_control(text = c("a", "b")), "`text` must be a single value")
  expect_error(wb$add_form_control(checked = "yes"), "`checked` must be TRUE or FALSE")
  expect_error(wb$set_page_setup(fit_to_width = c(1, 2)), "`fit_to_width` must be a single value")

  expect_silent(wb$set_page_setup(fit_to_width = 1, fit_to_height = 2))

})

test_that("values written as attributes are checked against their range", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$set_sheetview(zoom_scale = 409), "zoom_scale")
  expect_error(wb$set_sheetview(show_grid_lines = "yes"), "show_grid_lines")
  expect_error(wb$set_sheetview(view = "fancy"))
  expect_error(wb$page_setup(left = 49), "left")
  expect_error(wb$page_setup(fit_to_height = Inf), "fit_to_height")
  expect_error(wb$add_worksheet(hdpi = 0), "hdpi")
  expect_error(wb$add_worksheet(paper_size = 0), "paper_size")
  expect_error(wb$add_cell_style(dims = "A1", indent = 256), "indent")
  expect_error(wb$add_cell_style(dims = "A1", text_rotation = 200), "text_rotation")
  expect_error(wb$add_cell_style(dims = "A1", wrap_text = "yes"), "wrap_text")
  expect_error(wb$add_cell_style(dims = "A1", horizontal = "above"))
  expect_error(wb$set_row_heights(rows = 0, heights = 10), "rows")
  expect_error(wb$set_row_heights(rows = 1, heights = Inf), "heights")
  expect_error(wb$set_col_widths(cols = 1, hidden = NA), "hidden")
  expect_error(wb$set_base_font(font_size = 0), "font_size")
  expect_error(wb$set_base_font(font_name = "&#1;"), "font_name")
  expect_error(wb$add_font(dims = "A1", name = strrep("n", 32)), "31")
  expect_error(wb$add_font(dims = "A1", size = 410), "sz")
  expect_error(wb$add_font(dims = "A1", charset = 256), "charset")
  expect_error(wb$add_numfmt(dims = "A1", numfmt = -1), "numfmt")
  expect_error(wb_color(theme = 4, tint = 2), "tint")
  expect_error(wb_color(hex = "FFFFFFFFFF"), "valid color")
  expect_error(wb$add_named_region(dims = "A1", name = strrep("n", 256)), "255")
  expect_error(wb$add_named_region(dims = "A1", name = "x", hidden = "yes"), "hidden")
  expect_error(wb$add_hyperlink(dims = "A1", target = "https://example.com", tooltip = strrep("t", 256)), "255")
  expect_error(wb$set_header_footer(align_with_margins = NA), "align_with_margins")
  expect_error(create_sparklines("Sheet 1", "A2:A4", "B2", high = "yes"))
  expect_error(create_sparklines("Sheet 1", "A2:A4", "B2", line_weight = 1e10), "line_weight")

  expect_silent(wb$set_sheetview(zoom_scale = 150, show_grid_lines = FALSE, view = "pageLayout"))
  expect_silent(wb$page_setup(left = 1, fit_to_height = 1))
  expect_silent(wb$add_cell_style(dims = "A1", indent = 2, text_rotation = 255, horizontal = "center"))
  expect_silent(wb$set_row_heights(rows = 2, heights = 20, hidden = TRUE))
})
