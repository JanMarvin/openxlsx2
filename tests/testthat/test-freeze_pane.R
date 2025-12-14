test_that("Freeze Panes", {
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$freeze_pane(1, first_active_row = 3, first_active_col = 3)
  expected <- "<pane ySplit=\"2\" xSplit=\"2\" topLeftCell=\"C3\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$freeze_pane(1, first_active_row = 1, first_active_col = 3)
  expected <- "<pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$freeze_pane(1, first_active_row = 2, first_active_col = 1)
  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$freeze_pane(1, first_active_row = 2, first_active_col = 4)
  expected <- "<pane ySplit=\"1\" xSplit=\"3\" topLeftCell=\"D2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$freeze_pane(1, first_col = TRUE)
  expected <- "<pane xSplit=\"1\" topLeftCell=\"B1\" activePane=\"topRight\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$freeze_pane(1, first_row = TRUE)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$freeze_pane(1, first_row = TRUE, first_col = TRUE)

  expected <- "<pane ySplit=\"1\" xSplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)


  wb <- wb_workbook()
  # TODO can we add multiple at a time?
  wb$add_worksheet("Sheet 1")
  wb$add_worksheet("Sheet 2")
  wb$add_worksheet("Sheet 3")
  wb$add_worksheet("Sheet 4")
  wb$add_worksheet("Sheet 5")
  wb$add_worksheet("Sheet 6")
  wb$add_worksheet("Sheet 7")

  wb$freeze_pane(sheet = 1, first_active_row = 3, first_active_col = 3)
  wb$freeze_pane(sheet = 2, first_active_row = 1, first_active_col = 3)
  wb$freeze_pane(sheet = 3, first_active_row = 2, first_active_col = 1)
  wb$freeze_pane(sheet = 4, first_active_row = 2, first_active_col = 4)
  wb$freeze_pane(sheet = 5, first_col = TRUE)
  wb$freeze_pane(sheet = 6, first_row = TRUE)
  wb$freeze_pane(sheet = 7, first_row = TRUE, first_col = TRUE)


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

  # check that pre and post saving match
  temp <- temp_xlsx()
  on.exit(unlink(temp), add = TRUE)

  wb$save(temp)

  wb2 <- wb_load(temp)

  for (i in seq_along(wb$worksheets)) {
    expect_equal(
      wb$worksheets[[i]]$sheetViews,
      wb2$worksheets[[i]]$sheetViews
    )

    expect_equal(
      wb$worksheets[[i]]$freezePane,
      wb2$worksheets[[i]]$freezePane
    )
  }

})
