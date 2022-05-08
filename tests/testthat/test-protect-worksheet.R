test_that("Protection", {
  wb <- wb_workbook()
  wb$add_worksheet("s1")
  wb$add_worksheet("s2")

  wb <- wb_protect_worksheet(wb, sheet = "s1", protect = TRUE, password = "abcdefghij", properties =  c("formatCells", "formatColumns", "PivotTables"))
  expect_true(wb$worksheets[[1]]$sheetProtection == "<sheetProtection sheet=\"1\" formatCells=\"1\" formatColumns=\"1\" password=\"FEF1\"/>")

  wb <- wb_protect_worksheet(wb, sheet = "s2", protect = TRUE)
  expect_true(wb$worksheets[[2]]$sheetProtection == "<sheetProtection sheet=\"1\"/>")

  wb <- wb_protect_worksheet(wb, sheet = "s2", protect = FALSE)
  expect_equal(wb$worksheets[[2]]$sheetProtection, character())
})
