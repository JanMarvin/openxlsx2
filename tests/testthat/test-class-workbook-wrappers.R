
# wb_workbook() -----------------------------------------------------------

test_that("wb_workbook() is a wrapper", {
  ts <- Sys.time()
  expect_equal(wb_workbook(datetimeCreated = ts), wbWorkbook$new(datetimeCreated = ts))
  expect_wrapper("initialize", "wb_workbook", params = NULL)
})

# wb_add_worksheet() ------------------------------------------------------

test_that("wb_add_worksheet() is a wrapper", {
  expect_wrapper("addWorksheet", "wb_add_worksheet", params = list(sheet = "this"))
})

# wb_remove_worksheet() ---------------------------------------------------

test_that("wb_remove_worksheet() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  expect_wrapper("removeWorksheet", "wb_remove_worksheet", wb = wb, params = list(sheet = "sheet"))
})

# wb_save() ---------------------------------------------------------------

test_that("wb_save() is a wrapper", {
  # returns the file path instead
  expect_wrapper("save", "wb_save", params = NULL, ignore = "path")
})

# wb_merge_cells(), wb_unmerge_cells() ------------------------------------

test_that("wb_merge_cells(), wb_unmerge_cells() are wrappers", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  params <- list(sheet = "sheet", rows = 1:2, cols = 2)
  expect_wrapper("addCellMerge",    "wb_merge_cells",   wb = wb, params = params)
  expect_wrapper("removeCellMerge", "wb_unmerge_cells", wb = wb, params = params)
})

# wb_freeze_pane() --------------------------------------------------------

test_that("wb_freeze_pane() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  expect_wrapper("freezePanes", "wb_freeze_pane", wb = wb, params = list(sheet = "sheet"))
})

# wb_clone_worksheet() ----------------------------------------------------

test_that("wb_clone_worksheet() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  expect_wrapper("cloneWorksheet", "wb_clone_worksheet", wb = wb, params = list(old = "sheet", new = "new"))
})

# wb_add_style() ----------------------------------------------------------

test_that("wb_add_style() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  params <- list(sheet = "sheet", style = wb_style(), cols = 1, rows = 1)
  expect_wrapper("addStyle", "wb_add_style", wb = wb, params = params)
})

# wb_freeze_pane() --------------------------------------------------------

test_that("wb_freeze_pane() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  expect_wrapper("freezePanes", "wb_freeze_pane", wb = wb, params = list(sheet = "sheet"))
})

# wb_set_row_heights() --------------------------------------------------------

test_that("wb_set_row_heights() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  params <- list(sheet = "sheet", rows = 1, heights = 5)
  expect_wrapper("setRowHeights", "wb_set_row_heights", wb = wb, params = params)
})

# wb_group_rows() -------------------------------------------------------------

test_that("wb_group_rows() is a wrapper", {
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  params <- list(sheet = "sheet", rows = 1)
  expect_wrapper("groupRows", "wb_group_rows", wb = wb, params = params)
})

# wb_add_creators() -----------------------------------------------------------

test_that("wb_set_creators() is a wrapper", {
  expect_wrapper("addCreators", "wb_add_creators", params = list(creators = "myself"))
})

# wb_set_creators() -----------------------------------------------------------

test_that("wb_set_creators() is a wrapper", {
  expect_wrapper("setCreators", "wb_set_creators", params = list(creators = "myself"))
})

# wb_remove_creators() --------------------------------------------------------

test_that("wb_remove_creators() is a wrapper", {
  wb <- wb_workbook(creator = "myself")
  expect_wrapper("removeCreators", "wb_remove_creators", params = list(creators = "myself"))
})

