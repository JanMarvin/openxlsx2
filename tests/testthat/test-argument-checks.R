test_that("single value, flag and index arguments are checked", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$set_sheetview(zoom_scale = 1:3), "`zoom_scale` must be a single value.", fixed = TRUE)
  expect_error(wb$set_page_setup(copies = 1:3), "`copies` must be a single value.", fixed = TRUE)
  expect_error(wb$set_bookview(window_height = c(1, 2)), "`window_height` must be a single value.", fixed = TRUE)
  expect_error(wb$add_worksheet(zoom = c(100, 120)), "`zoom` must be a single value.", fixed = TRUE)

  expect_error(wb_add_data(wb, x = 1, start_col = 1.5), "`start_col` must be a single whole number.", fixed = TRUE)
  expect_error(wb_add_data(wb, x = 1, start_row = c(1, 2)), "`start_row` must be a single whole number.", fixed = TRUE)
  expect_error(wb_add_data(wb, x = 1, col_names = NA), "`col_names` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb_add_data(wb, x = 1, col_names = c(TRUE, FALSE)), "`col_names` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb_add_data_table(wb, x = data.frame(a = 1), with_filter = "yes"), "`with_filter` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb_add_formula(wb, x = "1+1", array = NA), "`array` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb_add_fill(wb, dims = 1), "`dims` must be a character vector of cell ranges like \"A1:B2\".", fixed = TRUE)
  expect_error(wb_add_font(wb, dims = NA_character_), "`dims` must be a character vector of cell ranges like \"A1:B2\".", fixed = TRUE)
  expect_error(wb_add_data(wb, x = 1, sheet = mean), "`sheet` must be a sheet name or a whole number sheet index.", fixed = TRUE)

  expect_error(wb_add_formula(wb, x = "1+1", array = 2), "`array` must be TRUE or FALSE.", fixed = TRUE)

  expect_silent(wb_add_data(wb, x = 1, start_col = "B"))
  expect_silent(wb_add_formula(wb, x = "1+1", array = 1, cm = 0L))
  expect_silent(wb$set_sheetview(zoom_scale = 120))
  expect_silent(wb$set_col_widths(cols = 1:2, widths = c(10, 20), hidden = c(TRUE, FALSE)))

})

test_that("cell style, ignore error and protection arguments are checked", {

  wb <- wb_workbook()$add_worksheet()$add_data(x = 1)

  expect_error(wb$add_cell_style(dims = "A1", wrap_text = c("1", "0")), "`wrap_text` must be a single value.", fixed = TRUE)
  expect_error(wb$add_ignore_error(dims = "A1", formula = "yes"), "`formula` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$protect_worksheet(protect = NA), "`protect` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$protect_worksheet(password = 1234), "`password` must be a single string.", fixed = TRUE)
  expect_error(wb$protect(lock_structure = "yes"), "`lock_structure` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$protect(password = c("a", "b")), "`password` must be a single string.", fixed = TRUE)

  expect_silent(wb$add_ignore_error(dims = "A1", number_stored_as_text = 1))
  expect_silent(wb$protect_worksheet(password = "pw"))

})

test_that("image, form control and page setup arguments are checked", {

  wb <- wb_workbook()$add_worksheet()
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  expect_error(wb$add_image(file = img, width = c(1, 2)), "`width` must be a single value.", fixed = TRUE)
  expect_error(wb$add_form_control(text = c("a", "b")), "`text` must be a single value.", fixed = TRUE)
  expect_error(wb$add_form_control(checked = "yes"), "`checked` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$set_page_setup(fit_to_width = c(1, 2)), "`fit_to_width` must be a single value.", fixed = TRUE)

  expect_silent(wb$set_page_setup(fit_to_width = 1, fit_to_height = 2))

})

test_that("numeric and enumerated arguments are checked against their range", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$set_sheetview(zoom_scale = 409), "`zoom_scale` must be a whole number from 10 to 400.", fixed = TRUE)
  expect_error(wb$set_sheetview(show_grid_lines = "yes"), "`show_grid_lines` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$set_sheetview(view = "fancy"), "x must be one of: 'normal', 'pageBreakPreview', 'pageLayout'", fixed = TRUE)
  expect_error(wb$page_setup(left = 49), "`left` must be a number from 0 to below 49.", fixed = TRUE)
  expect_error(wb$page_setup(fit_to_height = Inf), "`fit_to_height` must be a whole number from 0 to 32767.", fixed = TRUE)
  expect_error(wb$add_worksheet(hdpi = 0), "hdpi and vdpi must be positive", fixed = TRUE)
  expect_error(wb$add_worksheet(paper_size = 0), "paper_size must be an integer in range [1, 118]", fixed = TRUE)
  expect_error(wb$add_cell_style(dims = "A1", indent = 256), "`indent` must be a whole number from 0 to 255.", fixed = TRUE)
  expect_error(wb$add_cell_style(dims = "A1", text_rotation = 200), "`text_rotation` must be from 0 to 180, or 255 for vertical text.", fixed = TRUE)
  expect_error(wb$add_cell_style(dims = "A1", wrap_text = "yes"), "`wrap_text` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$add_cell_style(dims = "A1", horizontal = "above"), "x must be one of: 'general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'", fixed = TRUE)
  expect_error(wb$set_row_heights(rows = 0, heights = 10), "`rows` must be whole numbers from 1 to 1048576", fixed = TRUE)
  expect_error(wb$set_row_heights(rows = 1, heights = Inf), "`heights` must be finite numbers of at least 0", fixed = TRUE)
  expect_error(wb$set_col_widths(cols = 1, hidden = NA), "`hidden` must be TRUE or FALSE", fixed = TRUE)
  expect_error(wb$set_base_font(font_size = 0), "Invalid font_size, it must be a number from 1 to 409.55", fixed = TRUE)
  expect_error(wb$set_base_font(font_name = "&#1;"), "Invalid font_name, it must be a non-empty string with at most 31 characters", fixed = TRUE)
  expect_error(wb$add_font(dims = "A1", name = strrep("n", 32)), "`name` must be a non-empty string with at most 31 characters.", fixed = TRUE)
  expect_error(wb$add_font(dims = "A1", size = 410), "`sz` must be a number from 1 to 409.55.", fixed = TRUE)
  expect_error(wb$add_font(dims = "A1", charset = 256), "`charset` must be a whole number from 0 to 255.", fixed = TRUE)
  expect_error(wb$add_numfmt(dims = "A1", numfmt = -1), "a numeric `numfmt` must be a non-negative whole number", fixed = TRUE)
  expect_error(wb_color(theme = 4, tint = 2), "`tint` must be a number from -1 to 1", fixed = TRUE)
  expect_error(wb_color(hex = "FFFFFFFFFF"), "`color` ['FFFFFFFFFF'] is not a valid color", fixed = TRUE)
  expect_error(wb$add_named_region(dims = "A1", name = strrep("n", 256)), "`name` must have at most 255 characters", fixed = TRUE)
  expect_error(wb$add_named_region(dims = "A1", name = "x", hidden = "yes"), "`hidden` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$add_hyperlink(dims = "A1", target = "https://example.com", tooltip = strrep("t", 256)), "hyperlink tooltips are limited to 255 characters", fixed = TRUE)
  expect_error(wb$set_header_footer(align_with_margins = NA), "`align_with_margins` must be TRUE or FALSE", fixed = TRUE)
  expect_error(create_sparklines("Sheet 1", "A2:A4", "B2", high = "yes"), "high must be \"0\" or \"1\"", fixed = TRUE)
  expect_error(create_sparklines("Sheet 1", "A2:A4", "B2", line_weight = 1e10), "`line_weight` must be a number from 0 to 1584", fixed = TRUE)

  expect_silent(wb$set_sheetview(zoom_scale = 150, show_grid_lines = FALSE, view = "pageLayout"))
  expect_silent(wb$page_setup(left = 1, fit_to_height = 1))
  expect_silent(wb$add_cell_style(dims = "A1", indent = 2, text_rotation = 255, horizontal = "center"))
  expect_silent(wb$set_row_heights(rows = 2, heights = 20, hidden = TRUE))
})

test_that("wb_add_encharter() checks dims", {
  skip_if_not_installed("encharter")
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = data.frame(a = 1:3))
  ch <- encharter::ec("barChart")$add_series(data = "data!$A$2:$A$4")
  expect_error(wb_add_encharter(wb, graph = ch, dims = "A0"), "Row exceeds valid range", fixed = TRUE)
  expect_silent(wb_add_encharter(wb, graph = ch, dims = "C2:H10"))
})

test_that("bookview and sheet visibility arguments are checked", {
  wb <- wb_workbook()$add_worksheet()$add_worksheet()

  expect_error(wb$set_bookview(active_tab = 32767), "`active_tab` must be a whole number from 0 to 32766.", fixed = TRUE)
  expect_error(wb$set_bookview(window_width = 0.5), "`window_width` must be a whole number from 0 to 2147483647.", fixed = TRUE)
  expect_error(wb$set_bookview(minimized = "yes"), "`minimized` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$set_bookview(visibility = TRUE), "visibility must be one of \"visible\", \"hidden\" or \"veryHidden\"", fixed = TRUE)
  expect_error(wb$set_sheet_visibility(1, "bogus"), "`value` must be one of \"visible\", \"hidden\", \"veryHidden\", TRUE or FALSE", fixed = TRUE)
})

test_that("image arguments are checked", {
  wb <- wb_workbook()$add_worksheet()
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  expect_error(wb$add_image(file = img, width = NA), "`width` must be a number from 0 to Inf.", fixed = TRUE)
  expect_error(wb$add_image(file = img, height = 1e10), "`width` and `height` must not exceed 2348 inches.", fixed = TRUE)
  expect_error(wb$add_image(file = img, dims = "A0"), "`dims` must be a cell or a cell range like \"A1:B2\".", fixed = TRUE)
  expect_error(wb$add_image(file = img, col_offset = 0.5), "`col_offset` must be a whole number from -27273042329600 to 27273042316900.", fixed = TRUE)
})

test_that("page break and grouping arguments are checked", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$add_page_break(row = 0), "`row` must be a whole number from 1 to 1048576", fixed = TRUE)
  expect_error(wb$add_page_break(col = 16385), "Column exceeds valid range", fixed = TRUE)
  expect_error(wb$group_rows(rows = 1048577), "rows must be whole numbers from 1 to 1048576.", fixed = TRUE)
  expect_error(wb$group_rows(rows = 1:2, levels = c(8, 8)), "levels must be whole numbers from 0 to 7.", fixed = TRUE)
  expect_error(wb$group_cols(cols = 1:2, levels = c(0.5, 1)), "levels must be whole numbers from 0 to 7.", fixed = TRUE)
})

test_that("person and protection arguments are checked", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$add_person(id = "TRUE"), "id must be a guid in the form {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} with upper case hex digits", fixed = TRUE)
  expect_error(wb$add_person(name = ""), "name must be a non-empty string", fixed = TRUE)
  expect_error(wb$protect(file_sharing = TRUE, username = strrep("x", 55)), "`username` must be a non-empty string with at most 54 characters.", fixed = TRUE)
})

test_that("create functions check their arguments", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(create_numfmt(numFmtId = -1), "`numFmtId` must be a whole number from 0 to 4294967295.", fixed = TRUE)
  expect_error(create_numfmt(formatCode = strrep("0", 256)), "`formatCode` must be a non-empty string with less than 255 characters.", fixed = TRUE)
  expect_error(create_border(diagonal_up = "yes"), "`diagonal_up` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(create_cell_style(num_fmt_id = 1e10), "`num_fmt_id` must be a whole number from 0 to 4294967295.", fixed = TRUE)
  expect_error(create_cell_style(reading_order = 0.5), "`reading_order` must be a whole number from 0 to 2.", fixed = TRUE)
  expect_error(create_cell_style(text_rotation = 181), "`text_rotation` must be a whole number from 0 to 180 or 255.", fixed = TRUE)
  expect_error(create_font(name = "\uFFFF"), "`name` must be a non-empty string with at most 31 characters.", fixed = TRUE)
  expect_error(wb$add_data_validation(dims = "A1", type = "list", value = "\"a,b\"", prompt_title = strrep("t", 33)), "`prompt_title` must be a string with at most 32 characters", fixed = TRUE)
  expect_error(wb$add_conditional_formatting(dims = "A1:A3", type = "iconSet", rule = 1), "When type == 'iconSet', rule must contain at least two values", fixed = TRUE)
  expect_match(create_numfmt(numFmtId = 1e6), "numFmtId=\"1000000\"")
  expect_error(create_sparklines("Sheet 1", "A2:A4", "B2", color_series = wb_color(auto = TRUE)), "`color_series` must not be an automatic color", fixed = TRUE)

})

test_that("threads check person, comment and resolve", {
  wb <- wb_workbook()$add_worksheet()

  wb$add_person(name = "me")
  pid <- wb$get_person()$id[1]
  expect_error(wb$add_thread(dims = "B2", person_id = "me", comment = "x"), "person_id must be the id of a person, see wb_get_person()", fixed = TRUE)
  expect_error(wb$add_thread(dims = "B2", person_id = pid, resolve = TRUE), "no thread found in dims, comment is required", fixed = TRUE)
  expect_warning(wb$add_thread(dims = "B2", person_id = pid, comment = "x", reply = TRUE), "cannot reply")
  expect_silent(wb$add_thread(dims = "B2", person_id = pid, resolve = TRUE))

})

test_that("selected sheets, named region comments and chartsheets are saved", {
  wb <- wb_workbook()$add_worksheet()$add_worksheet()$add_data(x = 1)

  wb$set_selected(1:2)
  expect_equal(
    vapply(wb$worksheets, function(ws) grepl("tabSelected=\"1\"", ws$sheetViews), NA),
    c(TRUE, TRUE)
  )

  wb$add_named_region(sheet = 1, dims = "A1", name = "x", comment = "<b>&</b>")
  expect_silent(wb$save(temp_xlsx()))

  wb$add_chartsheet(visible = "veryhidden")
  expect_equal(wb$get_sheet_visibility()[3], "veryHidden")
})

test_that("pivot table, slicer and timeline params are checked", {
  df <- data.frame(a = c("x", "y"), b = 1:2, d = as.Date("2024-01-01") + 0:1)
  wb <- wb_workbook()$add_worksheet()$add_data(x = df)
  x <- wb_data(wb)

  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", params = list(row_grand_totals = "yes")), "`params$row_grand_totals` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", fun = "median"), "`fun` must be one of \"average\", \"count\", \"countNums\", \"max\", \"min\", \"product\", \"stdDev\", \"stdDevp\", \"sum\", \"var\", \"varp\".", fixed = TRUE)
  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", pivot_table = ""), "the pivot table name must be a string of 1 to 255 characters.", fixed = TRUE)
  expect_error(wb$add_pivot_table(x, rows = "a", data = "b", dims = ""), "Unexpected blank strings in dims validation detected", fixed = TRUE)

  wb$add_pivot_table(x, rows = "a", data = "b", slicer = "a", timeline = "d", pivot_table = "pt")
  expect_error(wb$add_slicer(x, slicer = "a", pivot_table = "pt", params = list(sort_order = "bogus")), "`params$sort_order` must be one of \"ascending\", \"descending\".", fixed = TRUE)
  expect_error(wb$add_slicer(x, slicer = "a", pivot_table = "pt", params = list(column_count = 0)), "`params$column_count` must be a whole number from 1 to 20000.", fixed = TRUE)
  expect_error(wb$add_timeline(x, timeline = "d", pivot_table = "pt", params = list(show_header = "yes")), "`params$show_header` must be TRUE or FALSE.", fixed = TRUE)
  expect_error(wb$add_timeline(x, timeline = "d", pivot_table = "pt", params = list(level = NaN)), "`params$level` must be a whole number from 0 to 4294967295.", fixed = TRUE)
  expect_error(wb$add_slicer(x, slicer = "a", pivot_table = "pt", params = list(caption = "&#1;")), "`params$caption` must be a non-empty string.", fixed = TRUE)
  expect_error(wb$add_timeline(x, timeline = "d", pivot_table = "pt", params = list(caption = "")), "`params$caption` must be a non-empty string.", fixed = TRUE)
  expect_error(wb$set_properties(datetime_created = NA), "`datetime_created` must be a single date-time", fixed = TRUE)
  expect_error(wb$set_properties(datetime_modified = "yesterday"), "`datetime_modified` must be a single date-time", fixed = TRUE)
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
  expect_error(wb$remove_bookview(view = NA), "`view` must be the index of a workbook view", fixed = TRUE)
  wb$remove_bookview(view = 1)
  expect_equal(wb$workbook$bookViews, "<bookViews><workbookView/></bookViews>")
})

test_that("pivot source names, table styles and text run arguments are checked", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(a = 1:3, b = c("x", "y", "z")))
  x <- wb_data(wb, 1)
  names(x)[2] <- NA
  expect_error(wb$add_pivot_table(x, rows = "a", data = "a"), "every column of a pivot table source needs a name", fixed = TRUE)

  expect_error(create_tablestyle(name = "ts", whole_table = 2.75), "`whole_table` must be the index of a dxf style, a whole number from 0.", fixed = TRUE)
  expect_error(create_pivottablestyle(name = ""), "`name` must be a single non-empty string.", fixed = TRUE)
  expect_error(create_pivottablestyle(name = "&#1;"), "`name` must be a single non-empty string.", fixed = TRUE)
  expect_error(create_pivottablestyle(name = "pts", blank_row = NaN), "`blank_row` must be the index of a dxf style, a whole number from 0.", fixed = TRUE)
  expect_error(wb$add_style(create_tablestyle(name = "ts", whole_table = 5)), "the table style references a dxf style that does not exist in the workbook", fixed = TRUE)

  expect_error(wb$set_cell_style(dims = "A1", style = 72), "`style` must be the id of a cell style in the workbook.", fixed = TRUE)
  expect_error(wb$set_cell_style(dims = "A1", style = -1), "`style` must be the id of a cell style in the workbook.", fixed = TRUE)
  expect_error(wb$set_cell_style(dims = "R1C1", style = 0), "`dims` must be cells like \"A1\" or a range like \"A1:B2\".", fixed = TRUE)
  expect_silent(wb$set_cell_style(dims = "A1:B2", style = 0))

  expect_error(fmt_txt("x", size = 0), "`size` must be a number from 1 to 409.55", fixed = TRUE)
  expect_error(fmt_txt("x", charset = 256), "`charset` must be a whole number from 0 to 255", fixed = TRUE)
  expect_error(fmt_txt("x", font = strrep("f", 32)), "`font` must be a non-empty string with at most 31 characters", fixed = TRUE)
  expect_error(fmt_txt("x", vert_align = ""), "`vert_align` must be one of \"baseline\", \"superscript\", \"subscript\"", fixed = TRUE)

  expect_error(create_shape(rotation = Inf), "`rotation` must be a number of degrees from -35791 to 35791.", fixed = TRUE)
  expect_error(create_shape(fill_transparency = 255), "`fill_transparency` must be a number from 0 to 100.", fixed = TRUE)
  expect_error(create_numfmt(formatCode = ""), "`formatCode` must be a non-empty string with less than 255 characters.", fixed = TRUE)
  expect_error(wb$add_numfmt(dims = "A1", numfmt = ""), "`formatCode` must be a non-empty string with less than 255 characters.", fixed = TRUE)
  expect_error(wb$add_named_style(name = strrep("x", 256)), "`name` must be a non-empty string with at most 255 characters", fixed = TRUE)
  expect_error(wb$add_named_region(dims = "A1", name = "cmt", comment = strrep("c", 256)), "`comment` must have at most 255 characters", fixed = TRUE)
  expect_error(wb$protect(file_sharing = TRUE, username = "\uFFFE"), "`username` must be a non-empty string with at most 54 characters.", fixed = TRUE)
  expect_error(create_shape(line_color = wb_color(theme = 99)), "`line_color` must be a theme color from 0 to 16.", fixed = TRUE)
  expect_match(create_shape(rotation = 1e-10), "rot=\"0\"")
})

test_that("limits beyond the schema are checked", {
  wb <- wb_workbook()$add_worksheet()

  expect_error(wb$set_header_footer(header = c(strrep("h", 100), strrep("h", 100), strrep("h", 50))), "a header or footer must have at most 255 characters", fixed = TRUE)
  expect_silent(wb$set_header_footer(header = c(strrep("h", 100), strrep("h", 100), strrep("h", 49))))
  expect_error(wb$set_header_footer(footer = c(NA, paste(rep("[Page]", 127), collapse = ""), NA)), "a header or footer must have at most 255 characters", fixed = TRUE)

  dv <- function(...) wb$add_data_validation(dims = "A1", type = "list", value = "\"a,b\"", ...)
  expect_error(dv(error = strrep("e", 225)), "`error` must be a string with at most 224 characters", fixed = TRUE)
  expect_error(dv(prompt = strrep("p", 256)), "`prompt` must be a string with at most 255 characters", fixed = TRUE)
  expect_silent(dv(error = strrep("e", 224), prompt = strrep("p", 255)))

  expect_error(wb$set_bookview(window_height = 2147483648), "`window_height` must be a whole number from 0 to 2147483647.", fixed = TRUE)
  expect_error(wb$set_bookview(window_width = 2147483648), "`window_width` must be a whole number from 0 to 2147483647.", fixed = TRUE)

  expect_error(wb$protect(file_sharing = TRUE, username = c("a", "b")), "`username` must be a single value.", fixed = TRUE)

  expect_error(create_numfmt(formatCode = strrep("0", 255)), "`formatCode` must be a non-empty string with less than 255 characters.", fixed = TRUE)
  expect_silent(create_numfmt(formatCode = strrep("0", 254)))

  for (i in seq_len(1022)) wb$add_page_break(row = i)
  expect_error(wb$add_page_break(row = 1023), "a worksheet can have at most 1022 row breaks", fixed = TRUE)
  for (i in seq_len(1023)) wb$add_page_break(col = i)
  expect_error(wb$add_page_break(col = 1024), "a worksheet can have at most 1023 column breaks", fixed = TRUE)
})

test_that("removing a worksheet before saving keeps the default view", {
  wb <- wb_workbook()$add_worksheet()$add_worksheet()$add_worksheet()
  wb$remove_worksheet(2)
  expect_null(wb$workbook$bookViews)
  temp <- temp_xlsx()
  wb$save(temp)
  expect_match(wb$workbook$bookViews, "<bookViews><workbookView ")
})
