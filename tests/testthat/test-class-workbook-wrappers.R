
# wb_workbook() -----------------------------------------------------------

test_that("wb_workbook() is a wrapper", {
  ts <- Sys.time()
  expect_equal(wb_workbook(datetimeCreated = ts), wbWorkbook$new(datetimeCreated = ts))
  expect_wrapper("initialize", "wb_workbook", params = NULL)
})

# wb_add_worksheet() ------------------------------------------------------

test_that("wb_add_worksheet() is a wrapper", {
  expect_wrapper("add_worksheet", "wb_add_worksheet", params = list(sheet = "this"))
})

# wb_remove_worksheet() ---------------------------------------------------

test_that("wb_remove_worksheet() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("remove_worksheet", "wb_remove_worksheet", wb = wb, params = list(sheet = "sheet"))
})

# wb_save() ---------------------------------------------------------------

test_that("wb_save() is a wrapper", {
  # returns the file path instead
  expect_wrapper("save", "wb_save", params = NULL, ignore = "path")
})

# wb_merge_cells(), wb_unmerge_cells() ------------------------------------

test_that("wb_merge_cells(), wb_unmerge_cells() are wrappers", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  params <- list(sheet = "sheet", rows = 1:2, cols = 2)
  expect_wrapper("merge_cells",    "wb_merge_cells",   wb = wb, params = params)
  expect_wrapper("unmerge_cells", "wb_unmerge_cells", wb = wb, params = params)
})

# wb_freeze_pane() --------------------------------------------------------

test_that("wb_freeze_pane() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("freeze_pane", "wb_freeze_pane", wb = wb, params = list(sheet = "sheet"))
})

# wb_clone_worksheet() ----------------------------------------------------

test_that("wb_clone_worksheet() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("clone_worksheet", "wb_clone_worksheet", wb = wb, params = list(old = "sheet", new = "new"))
})

# wb_freeze_pane() --------------------------------------------------------

test_that("wb_freeze_pane() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  expect_wrapper("freeze_pane", "wb_freeze_pane", wb = wb, params = list(sheet = "sheet"))
})

# wb_set_row_heights() --------------------------------------------------------

test_that("wb_set_row_heights() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  params <- list(sheet = "sheet", rows = 1, heights = 5)
  expect_wrapper("set_row_heights", "wb_set_row_heights", wb = wb, params = params)
})

# wb_remove_row_heights() -----------------------------------------------------

test_that("wb_remove_row_heights() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  wb$add_worksheet("a")
  wb$set_row_heights("a", 1:3, 20)
  params <- list(sheet = "sheet", rows = 2)
  expect_wrapper("remove_row_heights", "wb_remove_row_heights", wb = wb, params = params)
})

# wb_group_rows() -------------------------------------------------------------

test_that("wb_group_rows() is a wrapper", {
  wb <- wbWorkbook$new()$add_worksheet("sheet")
  params <- list(sheet = "sheet", rows = 1)
  expect_wrapper("wb_group_rows", "wb_group_rows", wb = wb, params = params)
})

# wb_set_creators() -----------------------------------------------------------

test_that("wb_set_creators() is a wrapper", {
  expect_wrapper("set_creators", "wb_set_creators", params = list(creators = "myself"))
})

# wb_remove_creators() --------------------------------------------------------

test_that("wb_remove_creators() is a wrapper", {
  wb <- wb_workbook(creator = "myself")
  expect_wrapper("remove_creators", "wb_remove_creators", params = list(creators = "myself"))
})

# wb_set_base_font() ----------------------------------------------------------

test_that("wb_set_base_font() is a wrapper", {
  params <- list(fontColour = "red", fontSize = 20)
  expect_wrapper("set_base_font", "wb_set_base_font", params = params)
})

# wb_set_header_footer() ------------------------------------------------------

test_that("wb_set_header_footer() is a wrapper", {
  wb <- wb_workbook(creator = "myself")$add_worksheet("a")
  expect_wrapper("set_header_footer", "wb_set_header_footer", wb = wb, params = list(sheet = "a"))

})

test_that("Workbook class wrappers work", {
  skip("no tests yet")
})

test_that("set_col_widths", {

  wb <- wbWorkbook$new()
  wb <- wb$add_worksheet("test")
  writeData(wb, "test", mtcars)

  # set column width to 12
  expect_silent(wb$set_col_widths("test", widths = 12L, cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" width=\"12\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # wrong sheet
  expect_error(wb$set_col_widths("test2", widths = 12L, cols = seq_along(mtcars)))

  # reset the column with, we do not provide an option ot remove the column entry
  expect_silent(wb$set_col_widths("test", cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" width=\"8.43\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # create column width for column 25
  expect_silent(wb$set_col_widths("test", cols = "Y", widths = 22))
  expect_equal(
    c("<col min=\"1\" max=\"11\" width=\"8.43\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>",
      "<col min=\"12\" max=\"24\" width=\"8.43\"/>",
      "<col min=\"25\" max=\"25\" width=\"22\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\"/>"),
    wb$worksheets[[1]]$cols_attr
  )

  # a few more errors
  expect_error(wb$set_col_widths("test", cols = "Y", width = 1:2))
  expect_error(wb$set_col_widths("test", cols = "Y", hidden = 1:2))
})

# wb_add_image() ----------------------------------------------------------

test_that("wb_add_image() is a wrapper", {
  path <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  wb <- wb_workbook()$add_worksheet("a")
  expect_wrapper("add_image", "wb_add_image", wb = wb, params = list(file = path, sheet = "a"))
})

# wb_add_plot() -----------------------------------------------------------

test_that("wb_add_plot() is a wrapper", {
  # plot is written to file. test can only be completed in interactive mode
  if(interactive()) {

    plot(1:5, 1:5)
    wb <- wb_workbook()$add_worksheet("a")

    # okay, not the best but the results have different field names.  Maybe that's
    # a feature to add to expect_wrapper()
    expect_error(
      openxlsx2:::expect_wrapper(
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
