
# wb_workbook() -----------------------------------------------------------

test_that("wb_workbook() is a wrapper", {
  ts <- Sys.time()
  expect_equal(wb_workbook(datetimeCreated = ts), wbWorkbook$new(datetimeCreated = ts))
  expect_wrapper("initialize", "wb_workbook", params = NULL)
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

# wb_save() ---------------------------------------------------------------

test_that("wb_save() is a wrapper", {
  # returns the file path instead
  expect_wrapper("save", params = NULL, ignore = "path")
})

# wb_merge_cells(), wb_unmerge_cells() ------------------------------------

test_that("wb_merge_cells(), wb_unmerge_cells() are wrappers", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  params <- list(sheet = "sheet", rows = 1:2, cols = 2)
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
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  wb$add_worksheet("a")
  wb$set_row_heights("a", 1:3, 20)
  params <- list(sheet = "sheet", rows = 2)
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

# wb_set_creators() -----------------------------------------------------------

test_that("wb_set_creators() is a wrapper", {
  expect_wrapper("set_creators", params = list(creators = "myself"))
})

# wb_remove_creators() --------------------------------------------------------

test_that("wb_remove_creators() is a wrapper", {
  wb <- wb_workbook(creator = "myself")
  expect_wrapper("remove_creators", params = list(creators = "myself"))
})

# wb_page_setup() -------------------------------------------------------------

test_that("wb_page_setup() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("page_setup", wb = wb, params = list(sheet = "a"))
})

# wb_set_base_font() ----------------------------------------------------------

test_that("wb_set_base_font() is a wrapper", {
  params <- list(fontColour = "red", fontSize = 20)
  expect_wrapper("set_base_font", params = params)
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
  # plot is written to file. test can only be completed in interactive mode
  if (interactive()) {

    plot(1:5, 1:5)
    wb <- wb_workbook()$add_worksheet("a")

    # okay, not the best but the results have different field names.  Maybe that's
    # a feature to add to expect_wrapper()
    expect_error(
      expect_wrapper(
        "add_plot",
        "wb_add_plot",
        wb = wb,
        params = list(sheet = "a")
      ),
      "wbWorkbook$add_plot$media$image1.png vs wb_add_plot$media$image1.png",
      fixed = TRUE
    )
  }
})

# wb_get_tables(), wb_remove_tables() -------------------------------------

test_that("wb_get_tables(), wb_remove_tables() are wrappers", {
  wb <- wb_workbook()
  wb$add_worksheet(sheet = "Sheet 1")
  write_datatable(wb, sheet = "Sheet 1", x = iris)
  write_datatable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
  expect_wrapper("get_tables", wb = wb, params = list(sheet = 1))
  expect_wrapper("remove_tables", wb = wb, params = list(sheet = 1, table = "mtcars"))
})

# wb_add_filter(), wb_remove_filter() -------------------------------------

test_that("wb_add_filter(), wb_remove_filter() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = "a", col = 1, row = 1)
  expect_wrapper("add_filter", wb = wb, params = params)
  wb$add_filter("a", 1, 1)
  expect_wrapper("remove_filter", wb = wb, params = list(sheet = "a"))
})

# wb_grid_lines() ---------------------------------------------------------

test_that("wb_grid_lines() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("grid_lines", wb = wb, params = list(sheet = 1, show = TRUE))
})

# wb_add_named_region(), wb_remove_named_region() -------------------------

test_that("wb_add_named_region(), wb_remove_named_region() are wrappers", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = 1, cols = 1, rows = 1, name = "cool")
  expect_wrapper("add_named_region", wb = wb, params = params)
  # now add the named region so that we can remove it
  wb$add_named_region(sheet = 1, cols = 1, rows = 1, name = "cool")
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
  # debugonce(expect_wrapper)
  expect_wrapper("get_sheet_visibility", wb = wb)
  expect_wrapper("set_sheet_visibility", wb = wb, params = list(sheet = 2:3, value = c(FALSE, FALSE)))
})

# wb_add_data_validation() ------------------------------------------------

test_that("wb_add_data_validation() is a wrapper", {
  wb <- wb_workbook()$add_worksheet("a")
  params <- list(sheet = 1, cols = 1, rows = 1, type = "whole", operator = "between", value = c(0, 1))
  expect_wrapper("add_data_validation", wb = wb, params = params)
})
