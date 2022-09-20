test_that("sheet.default_name", {

  wb <- wb_workbook()$add_worksheet()

  exp <- "Sheet 1"
  got <- names(wb$get_sheet_names())
  expect_equal(exp, got)

  op <- options("openxlsx2.sheet.default_name" = "Tabelle")
  on.exit(options(op), add = TRUE)
  wb <- wb_workbook()$add_worksheet()

  exp <- "Tabelle1"
  got <- names(wb$get_sheet_names())
  expect_equal(exp, got)

})
