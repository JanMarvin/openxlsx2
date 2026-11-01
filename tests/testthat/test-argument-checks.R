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

test_that("wb_add_encharter() checks dims", {
  skip_if_not_installed("encharter")
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = data.frame(a = 1:3))
  ch <- encharter::ec("barChart")$add_series(data = "data!$A$2:$A$4")
  expect_error(wb_add_encharter(wb, graph = ch, dims = "A0"), "Row exceeds")
  expect_silent(wb_add_encharter(wb, graph = ch, dims = "C2:H10"))
})

test_that("values written to workbook parts are checked", {
  wb <- wb_workbook()$add_worksheet()$add_worksheet()$add_data(x = 1)
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  expect_error(wb$set_bookview(active_tab = 32767), "active_tab")
  expect_error(wb$set_bookview(window_width = 0.5), "window_width")
  expect_error(wb$set_bookview(minimized = "yes"), "minimized")
  expect_error(wb$set_bookview(visibility = TRUE), "visibility")
  expect_error(wb$set_sheet_visibility(1, "bogus"), "value")
  expect_error(wb$add_image(file = img, width = NA), "width")
  expect_error(wb$add_image(file = img, height = 1e10), "2348")
  expect_error(wb$add_image(file = img, dims = "A0"), "dims")
  expect_error(wb$add_image(file = img, col_offset = 0.5), "col_offset")
  expect_error(wb$add_page_break(row = 0), "row")
  expect_error(wb$add_page_break(col = 16385), "Column")
  expect_error(wb$group_rows(rows = 1048577), "rows")
  expect_error(wb$group_rows(rows = 1:2, levels = c(8, 8)), "levels")
  expect_error(wb$group_cols(cols = 1:2, levels = c(0.5, 1)), "levels")
  expect_error(wb$add_person(id = "TRUE"), "id")
  expect_error(wb$add_person(name = ""), "name")
  expect_error(wb$protect(file_sharing = TRUE, username = strrep("x", 55)), "username")
  expect_error(create_numfmt(numFmtId = -1), "numFmtId")
  expect_error(create_numfmt(formatCode = strrep("0", 256)), "formatCode")
  expect_error(create_border(diagonal_up = "yes"), "diagonal_up")
  expect_error(create_cell_style(num_fmt_id = 1e10), "num_fmt_id")
  expect_error(create_cell_style(reading_order = 0.5), "reading_order")
  expect_error(create_cell_style(text_rotation = 181), "text_rotation")
  expect_error(create_font(name = "\uFFFF"), "name")
  expect_error(wb$add_data_validation(dims = "A1", type = "list", value = "\"a,b\"", prompt_title = strrep("t", 33)), "prompt_title")
  expect_error(wb$add_conditional_formatting(dims = "A1:A3", type = "iconSet", rule = 1), "two values")
  expect_match(create_numfmt(numFmtId = 1e6), "numFmtId=\"1000000\"")
  expect_error(create_sparklines("Sheet 1", "A2:A4", "B2", color_series = wb_color(auto = TRUE)), "color_series")

  wb$add_person(name = "me")
  pid <- wb$get_person()$id[1]
  expect_error(wb$add_thread(dims = "B2", person_id = "me", comment = "x"), "person_id")
  expect_error(wb$add_thread(dims = "B2", person_id = pid, resolve = TRUE), "comment")
  expect_warning(wb$add_thread(dims = "B2", person_id = pid, comment = "x", reply = TRUE), "cannot reply")
  expect_silent(wb$add_thread(dims = "B2", person_id = pid, resolve = TRUE))

  wb$set_selected(1:2)
  expect_equal(
    vapply(wb$worksheets, function(ws) grepl("tabSelected=\"1\"", ws$sheetViews), NA),
    c(TRUE, TRUE)
  )

  wb$add_chartsheet(visible = "veryhidden")
  expect_equal(wb$get_sheet_visibility()[3], "veryHidden")

  wb$add_named_region(sheet = 1, dims = "A1", name = "x", comment = "<b>&</b>")
  expect_silent(wb$save(temp_xlsx()))
})

test_that("pivot table, slicer and timeline params are checked", {
  df <- data.frame(a = c("x", "y"), b = 1:2, d = as.Date("2024-01-01") + 0:1)
  wb <- wb_workbook()$add_worksheet()$add_data(x = df)
  x <- wb_data(wb)

  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", params = list(row_grand_totals = "yes")), "row_grand_totals")
  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", fun = "median"), "fun")
  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", pivot_table = ""), "name")
  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", dims = ""), "blank")

  wb$add_pivot_table(x, rows = "a", data = "b", slicer = "a", timeline = "d", pivot_table = "pt")
  expect_error(wb$add_slicer(x, slicer = "a", pivot_table = "pt", params = list(sort_order = "bogus")), "sort_order")
  expect_error(wb$add_slicer(x, slicer = "a", pivot_table = "pt", params = list(column_count = 0)), "column_count")
  expect_error(wb$add_timeline(x, timeline = "d", pivot_table = "pt", params = list(show_header = "yes")), "show_header")
  expect_error(wb$add_timeline(x, timeline = "d", pivot_table = "pt", params = list(level = NaN)), "level")
  expect_error(wb$add_slicer(x, slicer = "a", pivot_table = "pt", params = list(caption = "&#1;")), "caption")
  expect_error(wb$add_timeline(x, timeline = "d", pivot_table = "pt", params = list(caption = "")), "caption")
  expect_error(wb$set_properties(datetime_created = NA), "datetime_created")
  expect_error(wb$set_properties(datetime_modified = "yesterday"), "datetime_modified")
})

test_that("escaped pivot table names and noncharacters in cells", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(a = 1:3, b = c("x", "y", "z")))
  wb$add_pivot_table(wb_data(wb, 1), rows = "b", data = "a", pivot_table = "<b>&</b>", slicer = "b")
  expect_silent(wb$add_slicer(wb_data(wb, 1), slicer = "b", pivot_table = "<b>&</b>"))
  expect_match(wb$pivotTables[1], "name=\"&lt;b&gt;&amp;&lt;/b&gt;\"")

  skip_if_not(l10n_info()[["UTF-8"]])
  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(s = c("a\uFFFE", " b\uFFFF ")))
  expect_equal(
    wb$worksheets[[1]]$sheet_data$cc$is[2:3],
    c("<is><t>a</t></is>", "<is><t xml:space=\"preserve\"> b </t></is>")
  )
})

test_that("removing parts leaves no references behind", {
  df <- data.frame(a = 1:3, b = c("x", "y", "z"), d = as.Date("2026-01-01") + 0:2)

  wb <- wb_workbook()$add_worksheet()$add_data(x = df)
  wb$add_pivot_table(wb_data(wb, 1), rows = "b", data = "a", pivot_table = "pt1", slicer = "b")
  wb$add_slicer(wb_data(wb, 1), slicer = "b", pivot_table = "pt1", sheet = 2)
  wb$remove_slicer(sheet = 2)
  expect_null(wb$workbook$extLst)
  expect_false(any(grepl("Slicer_", wb$workbook$definedNames)))

  wb <- wb_workbook()$add_worksheet()$add_data(x = df)
  wb$add_pivot_table(wb_data(wb, 1), rows = "d", data = "a", pivot_table = "pt1", timeline = "d")
  wb$add_timeline(wb_data(wb, 1), timeline = "d", pivot_table = "pt1", sheet = 2)
  wb$remove_timeline(sheet = 2)
  expect_null(wb$workbook$extLst)

  wb <- wb_workbook()$add_worksheet()$add_data(x = df)
  wb$add_pivot_table(wb_data(wb, 1), rows = "b", data = "a", pivot_table = "pt1")
  wb$remove_worksheet(2)
  expect_null(wb$workbook$pivotCaches)

  wb <- wb_workbook()$add_worksheet()$add_data(x = df)$
    add_comment(dims = "B2", comment = "a")$add_comment(dims = "C3", comment = "b")
  wb$remove_comment(dims = "A1:C3")
  expect_length(wb$vml, 0)

  wb <- wb_workbook()$add_worksheet()
  wb$set_bookview(active_tab = 0)
  expect_error(wb$remove_bookview(view = NA), "view")
  wb$remove_bookview(view = 1)
  expect_null(wb$workbook$bookViews)
})

test_that("values found by fuzzing the create functions are checked", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(a = 1:3, b = c("x", "y", "z")))
  x <- wb_data(wb, 1)
  names(x)[2] <- NA
  expect_error(wb$add_pivot_table(x, rows = "a", data = "a"), "needs a name")

  expect_error(create_tablestyle(name = "ts", whole_table = 2.75), "whole_table")
  expect_error(create_pivottablestyle(name = ""), "name")
  expect_error(create_pivottablestyle(name = "&#1;"), "name")
  expect_error(create_pivottablestyle(name = "pts", blank_row = NaN), "blank_row")
  expect_error(wb$add_style(create_tablestyle(name = "ts", whole_table = 5)), "dxf")

  expect_error(wb$set_cell_style(dims = "A1", style = 72), "style")
  expect_error(wb$set_cell_style(dims = "A1", style = -1), "style")
  expect_error(wb$set_cell_style(dims = "R1C1", style = 0), "dims")
  expect_silent(wb$set_cell_style(dims = "A1:B2", style = 0))

  expect_error(fmt_txt("x", size = 0), "size")
  expect_error(fmt_txt("x", charset = 256), "charset")
  expect_error(fmt_txt("x", font = strrep("f", 32)), "font")
  expect_error(fmt_txt("x", vert_align = ""), "vert_align")

  expect_error(create_shape(rotation = Inf), "rotation")
  expect_error(create_shape(fill_transparency = 255), "fill_transparency")
  expect_error(create_numfmt(formatCode = ""), "formatCode")
  expect_error(wb$add_numfmt(dims = "A1", numfmt = ""), "formatCode")
  expect_error(wb$add_named_style(name = strrep("x", 256)), "name")
  expect_error(wb$add_named_region(dims = "A1", name = "cmt", comment = strrep("c", 256)), "comment")
  expect_error(wb$protect(file_sharing = TRUE, username = "\uFFFE"), "username")
  expect_error(create_shape(line_color = wb_color(theme = 99)), "line_color")
  expect_match(create_shape(rotation = 1e-10), "rot=\"0\"")
})
