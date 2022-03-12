

test_that("wrappers are wrappers", {

  # wb_workbook() ----
  ts <- Sys.time()
  expect_equal(wb_workbook(datetimeCreated = ts), wbWorkbook$new(datetimeCreated = ts))
  expect_wrapper("initialize", "wb_workbook", params = NULL)

  # wb_add_worksheet() ----
  expect_wrapper("addWorksheet", "wb_add_worksheet", params = list(sheetName = "this"))

  # wb_save() ----
  # returns the file path instead
  expect_wrapper("save", "wb_save", params = NULL, ignore = "path")

  # wb_merge_cells(), wb_unmerge_cells() ----
  wb <- wbWorkbook$new()$addWorksheet("sheet")
  params <- list(sheet = "sheet", rows = 1:2, cols = 2)
  expect_wrapper("addCellMerge",    "wb_merge_cells",   wb = wb, params = params)
  expect_wrapper("removeCellMerge", "wb_unmerge_cells", wb = wb, params = params)
})
