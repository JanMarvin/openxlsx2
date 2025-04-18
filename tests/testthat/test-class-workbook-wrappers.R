testsetup()

# wb_workbook() -----------------------------------------------------------

test_that("wb_workbook() is a wrapper", {
  # slightly different be
  ts <- Sys.time()
  expect_identical(formals(wbWorkbook$public_methods$initialize), formals(wb_workbook))
  expect_identical(
    wbWorkbook$new(datetime_created = ts),
    wb_workbook(datetime_created = ts)
  )
})

# wb_add_chartsheet() ------------------------------------------------------

test_that("wb_add_chartsheet() is a wrapper", {
  expect_wrapper("add_chartsheet", params = list(sheet = "this"))
})

# wb_add_worksheet() ------------------------------------------------------

test_that("wb_add_worksheet() is a wrapper", {
  expect_wrapper("add_worksheet", params = list(sheet = "this"))
})

# wb_remove_worksheet() ---------------------------------------------------

test_that("wb_remove_worksheet() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("remove_worksheet", wb = wb, params = list(sheet = "sheet"))
})


# wb_to_df() ---------------------------------------------------------------

# does not work as expected
test_that("wb_to_df() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = iris)
  expect_pseudo_wrapper("to_df")
})

# wb_load() ---------------------------------------------------------------

test_that("wb_load() is a wrapper", {
  tmp_xlsx <- temp_xlsx()
  wb_workbook()$add_worksheet()$add_data(x = iris)$save(tmp_xlsx)
  expect_wrapper("load", params = list(file = tmp_xlsx), ignore_wb = TRUE)
})

# wb_save() ---------------------------------------------------------------

test_that("wb_save() is a wrapper", {
  # returns the file path instead
  expect_wrapper("save", params = NULL, ignore = c("file", "path"))
})

# wb_merge_cells(), wb_unmerge_cells() ------------------------------------

test_that("wb_merge_cells(), wb_unmerge_cells() are wrappers", {
  wb <- wb_workbook()$add_worksheet("sheet")
  params <- list(sheet = "sheet", dims = wb_dims(rows = 1:2, cols = 2))
  expect_wrapper("merge_cells", wb = wb, params = params)
  expect_wrapper("unmerge_cells", wb = wb, params = params)
})

# wb_freeze_pane() --------------------------------------------------------

test_that("wb_freeze_pane() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("freeze_pane", wb = wb, params = list(sheet = "sheet"))
})

# wb_clone_worksheet() ----------------------------------------------------

test_that("wb_clone_worksheet() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("clone_worksheet", wb = wb, params = list(old = "sheet", new = "new"))
})

# wb_freeze_pane() --------------------------------------------------------

test_that("wb_freeze_pane() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("freeze_pane", wb = wb, params = list(sheet = "sheet"))
})

# wb_set_row_heights() --------------------------------------------------------

test_that("wb_set_row_heights() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  params <- list(sheet = "sheet", rows = 1, heights = 5)
  expect_wrapper("set_row_heights", wb = wb, params = params)
})

# wb_remove_row_heights() -----------------------------------------------------

test_that("wb_remove_row_heights() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("a")
  wb$set_row_heights("a", 1:3, 20)
  params <- list(sheet = "a", rows = 2)
  expect_wrapper("remove_row_heights", wb = wb, params = params)
})

# wb_group_rows(), wb_ungroup_rows() ------------------------------------------

test_that("wb_group_rows(), wb_ungroup_rows() are wrappers", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  params <- list(sheet = "sheet", rows = 1)
  expect_wrapper("group_rows", wb = wb, params = params)
  wb$group_rows(sheet = "sheet", rows = 1)
  expect_wrapper("ungroup_rows", wb = wb, params = params)
})

# wb_group_cols(), wb_ungroup_cols() ------------------------------------------

test_that("wb_group_cols(), wb_ungroup_cols() are wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  params <- list(sheet = "sheet", cols = 1)
  # TODO add step to actual create columns?
  suppressMessages(expect_wrapper("group_cols", wb = wb, params = params))
  suppressMessages(wb$group_cols(sheet = "sheet", cols = 1))
  expect_wrapper("ungroup_cols", wb = wb, params = params)
})

# wb_set_creators() ------------------------------------------------------------

test_that("wb_set_creators() is a wrapper", {
  expect_wrapper("set_creators", params = list(creators = "myself"))
})

# wb_remove_creators() ---------------------------------------------------------

test_that("wb_remove_creators() is a wrapper", {
  wb <- wb_workbook(creator = "myself")
  expect_wrapper("remove_creators", wb = wb, params = list(creators = "myself"))
})

# wb_set_properties() ----------------------------------------------------------

test_that("wb_set_properties() is a wrapper", {
  wb <- wb_workbook(subject = "xyz")
  expect_wrapper("set_properties", wb = wb, params = list(subject = "a new subject"))
})

# wb_set_properties() ---------------------------------------------------------

test_that("wb_get_properties() is a wrapper", {
  wb <- wb_workbook(subject = "xyz")
  expect_wrapper("get_properties", wb = wb)
})

# wb_add_mips() ---------------------------------------------------------------

test_that("wb_add_mips() is a wrapper", {
  wb <- wb_workbook()
  xml <- "<property fmtid=\"1\" pid=\"2\" name=\"MSIP_Label_1_Enabled\"><vt:lpwstr>true</vt:lpwstr></property>"
  expect_wrapper("add_mips", wb = wb, params = list(xml = xml))
})

# wb_get_mips() ---------------------------------------------------------------

test_that("wb_get_mips() is a wrapper", {
  xml <- "<property fmtid=\"1\" pid=\"2\" name=\"MSIP_Label_1_Enabled\"><vt:lpwstr>true</vt:lpwstr></property>"
  wb <- wb_workbook()$add_mips(xml = xml)
  expect_wrapper("get_mips", wb = wb)
})

# wb_set_last_modified_by() ---------------------------------------------------

test_that("wb_set_last_modified_by() is a wrapper", {
  wb <- wb_workbook()
  expect_wrapper("set_last_modified_by", params = list(name = "myself"))
})

# wb_page_setup() -------------------------------------------------------------

test_that("wb_page_setup() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("page_setup", wb = wb, params = list(sheet = "a"))
})

# wb_set_base_font() ----------------------------------------------------------

test_that("wb_set_base_font() is a wrapper", {
  params <- list(font_colour = "red", font_size = 20)
  expect_wrapper("set_base_font", params = params)
})

test_that("wb_get_base_font() is a wrapper", {
  wb <- wb_workbook()
  expect_wrapper("get_base_font", wb = wb)
  expect_wrapper("get_base_font", wb = wb)
})

test_that("wb_set_base_colors() is a wrapper", {
  params <- list(theme = 2)
  expect_wrapper("set_base_colors", params = params)
  expect_wrapper("set_base_colours", params = params)
})

test_that("wb_get_base_color() is a wrapper", {
  wb <- wb_workbook()
  expect_wrapper("get_base_colors", wb = wb)
  expect_wrapper("get_base_colours", wb = wb)
})

# wb_get_bookview(),wb_remove_bookview(), wb_set_bookview() -------------------

test_that("wb_get_bookview() is a wrapper", {
  expect_wrapper("get_bookview")
})

test_that("wb_set_bookview() is a wrapper", {
  params <- list(active_tab = "1")
  expect_wrapper("set_bookview", params = params)
})

test_that("wb_remove_bookview() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()$set_bookview()
  params <- list(view = 1L)
  expect_wrapper("remove_bookview", params = params)
})

# wb_set_header_footer() ------------------------------------------------------

test_that("wb_set_header_footer() is a wrapper", {
  wb <- wb_workbook(creator = "myself")$add_worksheet("a")
  expect_wrapper("set_header_footer", wb = wb, params = list(sheet = "a"))
})

# wb_set_col_widths(), wb_remove_col_widths() -----------------------------

test_that("wb_set_col_widths() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("set_col_widths", wb = wb, params = list(sheet = "a", cols = 1:3))
  expect_wrapper("remove_col_widths", wb = wb, params = list(sheet = "a", cols = 1:3))
})

# wb_add_image() ----------------------------------------------------------

test_that("wb_add_image() is a wrapper", {
  path <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("add_image", wb = wb, params = list(file = path, sheet = "a"))
})

# wb_add_plot() -----------------------------------------------------------

test_that("wb_add_plot() is a wrapper", {

  # workaround: this filename is inserted to the wrapper function
  op <- options("openxlsx2.temp_png" = tempfile(pattern = "figureImage", fileext = ".png"))
  on.exit(options(op), add = TRUE)

  # create a device we can dev.copy() from
  grDevices::pdf(NULL) # do not create "Rplots.pdf"
  grDevices::dev.control("enable")
  plot(1:5, 1:5)

  wb <- wb_workbook()$add_worksheet("a")

  expect_wrapper(
    "add_plot",
    "wb_add_plot",
    wb = wb,
    params = list(sheet = "a")
  )

  # # check that it is actually working
  # wb$add_plot(sheet = "a")$save("~/test.xlsx")

  grDevices::dev.off()

})

test_that("wb_add_drawing is a wrapper", {

  fl <- testfile_path("loadExample.xlsx")
  wb <- wb_load(file = fl)

  xml <- wb$drawings[[2]]

  wb <- wb_workbook()$add_worksheet()

  expect_wrapper("add_drawing", wb = wb, params = list(xml = xml))
})

# wb_get_tables(), wb_remove_tables() -------------------------------------

test_that("wb_get_tables(), wb_remove_tables() are wrappers", {
  wb <- wb_workbook()
  wb$add_worksheet(sheet = "Sheet 1")
  wb$add_data_table(sheet = "Sheet 1", x = iris)
  wb$add_data_table(sheet = 1, x = mtcars, table_name = "mtcars", start_col = 10)
  expect_wrapper("get_tables", wb = wb, params = list(sheet = 1))
  expect_wrapper("remove_tables", wb = wb, params = list(sheet = 1, table = "mtcars"))
})

# wb_add_filter(), wb_remove_filter() -------------------------------------

test_that("wb_add_filter(), wb_remove_filter() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = "a", cols = 1, rows = 1)
  expect_wrapper("add_filter", wb = wb, params = params)
  wb$add_filter("a", 1, 1)
  expect_wrapper("remove_filter", wb = wb, params = list(sheet = "a"))
})

# wb_grid_lines() ---------------------------------------------------------

test_that("wb_set_grid_lines() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("set_grid_lines", wb = wb, params = list(sheet = 1, show = TRUE))
})

# wb_add_named_region(), wb_remove_named_region() -------------------------

test_that("wb_add_named_region(), wb_remove_named_region() wb_get_named_regions() are wrappers", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = 1, dims = "A1", name = "cool")
  expect_wrapper("add_named_region", wb = wb, params = params)
  # now add the named region so that we can remove it
  wb$add_named_region(sheet = 1, dims = "A1", name = "cool")
  expect_wrapper("get_named_regions", wb = wb)
  expect_wrapper("remove_named_region", wb = wb, params = list(name = "cool"))
})

# wb_set_order() ----------------------------------------------------------

test_that("wb_set_order() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")$add_worksheet("b")$add_worksheet("c")
  expect_wrapper("set_order", wb = wb, params = list(sheets = 3:1))
})

# wb_get_sheet_visibility(), wb_set_sheet_visibility() --------------------

test_that("wb_get_sheet_visibility(), wb_set_sheet_visibility() are wrappers", {
  wb <- wb_workbook()$add_worksheet("a")$add_worksheet("b")$add_worksheet("c")
  expect_wrapper("get_sheet_visibility", wb = wb)
  expect_wrapper("set_sheet_visibility", wb = wb, params = list(sheet = 2:3, value = c(FALSE, FALSE)))
})

# wb_add_data_validation() ------------------------------------------------

test_that("wb_add_data_validation() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = 1, dims = "A1", type = "whole", operator = "between", value = c(0, 1))
  expect_wrapper("add_data_validation", wb = wb, params = params)
})

# wb_protect(), wb_protect_worksheet() ------------------------------------

test_that("wb_protect() is a wrapper", {
  expect_wrapper("protect")
})

test_that("wb_protect_worksheet() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = "a", properties = c("deleteRows", "autoFilter"))
  expect_wrapper("protect_worksheet", wb = wb, params = params)
})

# wb_add_page_break() -----------------------------------------------------

test_that("wb_add_page_break() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("add_page_break", wb = wb, params = list(sheet = 1, row = 2))
})

# wb_clean_sheet() --------------------------------------------------------

test_that("wb_clean_sheet() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = "a")
  expect_wrapper("clean_sheet", wb = wb, params = params)
})

## disable test bailing in covr in GitHub action
# # wb_open() ---------------------------------------------------------------
#
# test_that("wb_open() is a wrapper", {
#   wb <- wb_workbook()$add_worksheet("a")
#   expect_error(
#     expect_wrapper("open", wb = wb),
#     "wbWorkbook$open$path vs wb_open$path",
#     fixed = TRUE
#   )
# })

# wb_add_data(), wb_add_data_table(), wb_add_formula() --------------------

test_that("wb_add_data() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper("add_data",       wb = wb, params = list(sheet = 1, x = data.frame(x = 1)))
})

test_that("wb_add_data_table() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper("add_data_table", wb = wb, params = list(sheet = 1, x = data.frame(x = 1)))
})

test_that("wb_add_pivot_table() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = mtcars)
  df <- wb_data(wb)
  expect_wrapper("add_pivot_table", wb = wb, params = list(x = df, data = "disp"))
})

test_that("wb_add_slicer() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = mtcars)
  df <- wb_data(wb)
  wb$add_pivot_table(x = df, data = "disp", slicer = "vs", pivot_table = "pivot1")
  expect_wrapper("add_slicer", wb = wb, params = list(x = df, slicer = "vs", pivot_table = "pivot1"))
})

test_that("wb_remove_slicer() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()
  expect_wrapper("remove_slicer", wb = wb)
})

test_that("wb_add_timeline() is a wrapper", {
  df <- data.frame(
    date = seq(from = as.Date("2024-01-01"), length.out = 26, by = "month"),
    amnt = sample(seq(100:150), 26, replace = TRUE)
  )

  wb <- wb_workbook()$add_worksheet()$add_data(x = df)
  df <- wb_data(wb)
  wb$add_pivot_table(x = df, data = "amnt", timeline = "date", pivot_table = "pivot1")
  expect_wrapper("add_timeline", wb = wb, params = list(x = df, timeline = "date", pivot_table = "pivot1"))
})

test_that("wb_remove_timeline() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()
  expect_wrapper("remove_timeline", wb = wb)
})

test_that("wb_add_formula() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()
  expect_wrapper("add_formula",    wb = wb, params = list(sheet = 1, x = "=TODAY()"))
})

test_that("wb_add_hyperlink() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = "http://github.com/JanMarvin/openxlsx2")
  expect_wrapper("add_hyperlink", wb = wb, params = list(sheet = 1, col_names = FALSE))
})

test_that("wb_remove_hyperlink() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()
  expect_wrapper("remove_hyperlink", wb = wb)
})

test_that("wb_update_table() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()$add_data_table(x = iris[1:10, ])
  expect_wrapper("update_table", wb = wb, params = list(sheet = 1, tabname = "Table1", dims = "A1:D4"))
})

test_that("wb_copy_cells() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)$add_data(x = "1")
  dat <- wb_data(wb, 1, dims = "A1", col_names = FALSE)
  expect_wrapper("copy_cells",     wb = wb, params = list(sheet = 1, data = dat, dims = "B1"))
})

# wb_add_comment() --------------------------------------------------------

test_that("wb_add_comment() is a wrapper", {

  c1 <- wb_comment(text = "this is a comment", author = "")

  wb <- wb_workbook()$add_worksheet()
  expect_wrapper(
    "add_comment",
    wb = wb,
    params = list(comment = c1, dims = "A1")
  )

  op <- options("openxlsx2.creator" = "user")
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()$add_worksheet()

  expect_wrapper(
    "add_comment",
    wb = wb,
    params = list(comment = "a new comment", dims = "A1")
  )

})

test_that("wb_get_comment() is a wrapper", {

  c1 <- wb_comment(text = "this is a comment", author = "")
  wb <- wb_workbook()$add_worksheet("Sheet 1")$add_comment(dims = "B10", comment = c1)

  expect_wrapper(
    "get_comment",
    wb = wb
  )

})

# wb_remove_comment() -----------------------------------------------------

test_that("wb_remove_comment() is a wrapper", {

  c1 <- wb_comment(text = "this is a comment", author = "")

  wb <- wb_workbook()$
    add_worksheet()$
    add_comment(dims = "A1", comment = c1)

  expect_wrapper(
    "remove_comment",
    wb = wb,
    params = list(dims = "A1")
  )

})

# wb_add_thread() ---------------------------------------------------------

test_that("wb_add_thread() is a wrapper", {

  wb <- wb_workbook()$add_worksheet()

  expect_wrapper(
    "add_person",
    wb = wb,
    params = list(name = "me")
  )

  wb <- wb_workbook()$add_worksheet()$add_person("me")

  expect_wrapper(
    "get_person",
    wb = wb,
    params = list(name = "me")
  )

  wb <- wb_workbook()$add_worksheet()$add_person("me")
  me_id <- wb$get_person("me")$id

  expect_wrapper(
    "add_thread",
    wb = wb,
    params = list(comment = "test", person_id = me_id)
  )

})

test_that("wb_get_thread() is a wrapper", {

  wb <- wb_workbook()$add_worksheet()$add_person("me")
  me_id <- wb$get_person("me")$id
  wb$add_thread(comment = "test", person_id = me_id)

  expect_wrapper(
    "get_thread",
    wb = wb
  )

})

# wb_add_form_control() ---------------------------------------------------

test_that("wb_add_form_control() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()
  expect_wrapper(
    "add_form_control",
    wb = wb
  )
})

# wb_add_conditional_formatting() -----------------------------------------

test_that("wb_add_conditional_formatting() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  params <- list(sheet = 1, dims = wb_dims(cols = 1, rows = 1), type = "topN")
  expect_wrapper(
    "add_conditional_formatting",
    wb = wb,
    params = params,
    ignore_fields = "styles_mgr"
  )
})

# wb_remove_conditional_formatting() -----------------------------------------

test_that("wb_remove_conditional_formatting() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()
  expect_wrapper(
    "remove_conditional_formatting",
    wb = wb
  )
})

# wb_set_sheet_names() ----------------------------------------------------

test_that("wb_set_sheet_names() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")$add_worksheet("b")
  expect_wrapper("set_sheet_names", wb = wb, params = list(new = c("c", "d")))
})

# wb_get_sheet_names() ----------------------------------------------------

test_that("wb_get_sheet_names() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")$add_worksheet("b")
  expect_wrapper("get_sheet_names", wb = wb)
})

# wb_clean_sheet() --------------------------------------------------------

test_that("wb_clean_sheet() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper("clean_sheet", wb = wb, params = list(sheet = 1))
})

# wb_add_border() ---------------------------------------------------------

test_that("wb_add_border() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper(
    "add_border",
    wb = wb,
    params = list(sheet = 1)
  )
})

# wb_add_fill() -----------------------------------------------------------

test_that("wb_add_fill() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper(
    "add_fill",
    wb = wb,
    params = list(sheet = 1)
  )
})

# wb_add_font() -----------------------------------------------------------

test_that("wb_add_font() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper(
    "add_font",
    wb = wb,
    params = list(sheet = 1)
  )
})

# wb_add_numfmt() ---------------------------------------------------------

test_that("wb_add_numfmt() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper(
    "add_numfmt",
    wb = wb,
    params = list(sheet = 1, numfmt = 1)
  )
})

# wb_add_cell_style() -----------------------------------------------------

test_that("wb_add_cell_style() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper(
    "add_cell_style",
    wb = wb,
    params = list(sheet = 1)
  )
})

# wb_add_named_style() ----------------------------------------------------

test_that("wb_add_named_style() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper(
    "add_named_style",
    wb = wb,
    params = list(name = "Title")
  )
})

# wb_add_dxfs_style() -----------------------------------------------------

test_that("wb_add_dxfs_style() is a wrapper", {
  wb <- wb_workbook()$add_worksheet(1)
  expect_wrapper(
    "add_dxfs_style",
    wb = wb,
    params = list(name = "dxfs")
  )
})

# wb_clone_sheet_style() --------------------------------------------------

test_that("wb_clone_sheet_style() is a wrapper", {
  wb <- wb_workbook()$
    add_worksheet(1)$add_data(x = mtcars)$
    add_worksheet(2)$add_data(x = mtcars)
  wb$add_fill(sheet = 1, dims = "D5:J23", color = wb_colour(hex = "FFFFFF00"))
  expect_wrapper(
    "clone_sheet_style",
    wb = wb,
    params = list(from = 1, to = 2)
  )
})


# wb_add_sparklines() -----------------------------------------------------

test_that("wb_add_sparklines() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = mtcars)
  expect_wrapper(
    "add_sparklines",
    wb = wb,
    params = list(sparklines = create_sparklines("Sheet 1", "A3:L3", "M3", type = "column", first = "1"))
  )
})


# wb_set_sheetview() ------------------------------------------------------

test_that("wb_set_sheetview() is a wrapper", {
  wb <- wb_workbook()$add_worksheet()
  expect_wrapper(
    "set_sheetview",
    wb = wb,
    params = list(right_to_left = TRUE)
  )
})


# wb_add_style() ----------------------------------------------------------

test_that("wb_add_style() is a wrapper", {

  # with name
  style <- create_numfmt(numFmtId = "165", formatCode = "#.#")
  wb <- wb_workbook()
  expect_wrapper(
    "add_style",
    wb = wb,
    params = list(style = style, style_name = "numfmt")
  )

  # without name
  wb <- wb_workbook()
  numfmt <- create_numfmt(numFmtId = "165", formatCode = "#.#")
  expect_wrapper(
    "add_style",
    wb = wb,
    params = list(style = numfmt)
  )

})


# wb_set_cell_style_across() ----------------------------------------------

test_that("wb_add_style() is a wrapper", {

  wb <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_fill(dims = "C3", color = wb_color("yellow"))

  expect_wrapper(
    "set_cell_style_across",
    wb = wb,
    params = list(style = "C3", cols = "C:D", rows = 3:4)
  )

})


# wb_get_cell_style() -----------------------------------------------------

test_that("wb_get_cell_style() is a wrapper", {

  # set a style in A1
  wb <- wb_workbook()$add_worksheet()$add_numfmt(dims = "A1", numfmt = "#.0")

  expect_wrapper(
    "get_cell_style",
    wb = wb,
    params = list(dims = "A1")
  )

})


# wb_set_cell_style() -----------------------------------------------------

test_that("wb_set_cell_style() is a wrapper", {

  # set a style in b1
  wb <- wb_workbook()$add_worksheet()$
    add_numfmt(dims = "B1", numfmt = "#,0")

  # get style from b1 to assign it to a1
  numfmt <- wb$get_cell_style(dims = "B1")

  expect_wrapper(
    "set_cell_style",
    wb = wb,
    params = list(dims = "A1", style = numfmt)
  )

})


# wb_add_chart_xml() ------------------------------------------------------

test_that("wb_add_chart_xml() is a wrapper", {

  wb <- wb_workbook()$add_worksheet()

  expect_wrapper(
    "add_chart_xml",
    wb = wb,
    params = list(dims = "F4:L20", xml = "<a/>")
  )
})


# wb_add_mschart() --------------------------------------------------------

test_that("wb_add_mschart() is a wrapper", {

  skip_if_not_installed("mschart")

  require(mschart)

  ### Scatter
  scatter <- ms_scatterchart(data = iris, x = "Sepal.Length",
                             y = "Sepal.Width", group = "Species")
  scatter <- chart_settings(scatter, scatterstyle = "marker")

  wb <- wb_workbook()$add_worksheet()

  expect_wrapper(
    "add_mschart",
    wb = wb,
    params = list(dims = "F4:L20", graph = scatter)
  )

})

# wb_add_ignore_error() ---------------------------------------------------

test_that("add_ignore_error is a wrapper", {

  wb <- wb_workbook()$add_worksheet()

  expect_wrapper(
    "add_ignore_error",
    wb = wb,
    params = list(dims = "F4:L20", formula = TRUE)
  )

})
