

test_that("Freeze Panes", {
  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$freezePanes(1, firstActiveRow = 3, firstActiveCol = 3)
  expected <- "<pane ySplit=\"2\" xSplit=\"2\" topLeftCell=\"C3\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$freezePanes(1, firstActiveRow = 1, firstActiveCol = 3)
  expected <- "<pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$freezePanes(1, firstActiveRow = 2, firstActiveCol = 1)
  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$freezePanes(1, firstActiveRow = 2, firstActiveCol = 4)
  expected <- "<pane ySplit=\"1\" xSplit=\"3\" topLeftCell=\"D2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$freezePanes(1, firstCol = TRUE)
  expected <- "<pane xSplit=\"1\" topLeftCell=\"B1\" activePane=\"topRight\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$freezePanes(1, firstRow = TRUE)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$addWorksheet("Sheet 1")
  wb$freezePanes(1, firstRow = TRUE, firstCol = TRUE)

  expected <- "<pane ySplit=\"1\" xSplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)


  wb <- wb_workbook()
  # TODO can we add multiple at a time?
  wb$addWorksheet("Sheet 1")
  wb$addWorksheet("Sheet 2")
  wb$addWorksheet("Sheet 3")
  wb$addWorksheet("Sheet 4")
  wb$addWorksheet("Sheet 5")
  wb$addWorksheet("Sheet 6")
  wb$addWorksheet("Sheet 7")

  wb$freezePanes(sheet = 1, firstActiveRow = 3, firstActiveCol = 3)
  wb$freezePanes(sheet = 2, firstActiveRow = 1, firstActiveCol = 3)
  wb$freezePanes(sheet = 3, firstActiveRow = 2, firstActiveCol = 1)
  wb$freezePanes(sheet = 4, firstActiveRow = 2, firstActiveCol = 4)
  wb$freezePanes(sheet = 5, firstCol = TRUE)
  wb$freezePanes(sheet = 6, firstRow = TRUE)
  wb$freezePanes(sheet = 7, firstRow = TRUE, firstCol = TRUE)


  expected <- "<pane ySplit=\"2\" xSplit=\"2\" topLeftCell=\"C3\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  expected <- "<pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\"/>"
  expect_equal(wb$worksheets[[2]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/>"
  expect_equal(wb$worksheets[[3]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" xSplit=\"3\" topLeftCell=\"D2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[4]]$freezePane, expected)

  expected <- "<pane xSplit=\"1\" topLeftCell=\"B1\" activePane=\"topRight\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[5]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[6]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" xSplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[7]]$freezePane, expected)
})
