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

test_that("logical vector works too", {

  wb <- wb_workbook()$add_worksheet("S1")

  # pre 0.7.1 style
  wb$protect_worksheet(
    "S1",
    protect = TRUE,
    properties = c("formatCells", "formatColumns", "insertColumns", "deleteColumns")
  )

  exp <- "<sheetProtection sheet=\"1\" formatCells=\"1\" formatColumns=\"1\" insertColumns=\"1\" deleteColumns=\"1\"/>"
  got <- wb$worksheets[[1]]$sheetProtection
  expect_equal(exp, got)

  # post 0.7 style
  wb$protect_worksheet(
    "S1",
    protect = TRUE,
    properties = c(formatCells = FALSE, formatColumns = FALSE, insertColumns = TRUE, deleteColumns = TRUE)
  )

  exp <- "<sheetProtection sheet=\"1\" formatCells=\"0\" formatColumns=\"0\" insertColumns=\"1\" deleteColumns=\"1\"/>"
  got <- wb$worksheets[[1]]$sheetProtection
  expect_equal(exp, got)

})
