test_that("Page setup", {
  wb <- wb_workbook()
  wb$add_worksheet("s1")
  wb$add_worksheet("s2")

  wb$page_setup(
    sheet = "s1", orientation = "landscape", scale = 100, left = 0.1,
    right = 0.1, top = .75, bottom = .75, header = 0.1, footer = 0.1,
    fitToWidth = TRUE, fitToHeight = TRUE, paperSize = 1,
    summaryRow = "below", summaryCol = "right"
  )

  wb$page_setup(
    sheet = 2, orientation = "landscape", scale = 100, left = 0.1,
    right = 0.1, top = .75, bottom = .75, header = 0.1, footer = 0.1,
    fitToWidth = TRUE, fitToHeight = TRUE, paperSize = 1,
    summaryRow = "below", summaryCol = "right"
  )

  expect_equal(wb$worksheets[[1]]$pageSetup, wb$worksheets[[2]]$pageSetup)

  v <- gsub(" ", "", wb$worksheets[[1]]$pageSetup, fixed = TRUE)
  expect_true(grepl('paperSize="1"', v))
  expect_true(grepl('orientation="landscape"', v))
  expect_true(grepl('fitToWidth="1"', v))
  expect_true(grepl('fitToHeight="1"', v))

  pr <- wb$worksheets[[1]]$sheetPr

  # SheetPr will be a character vector of length 2; the first entry will
  # be for pageSetupPr, inserted by `fitToWidth`/`fitToHeight`.
  expect_true(grepl('<outlinePr summaryBelow="1" summaryRight="1"/>', pr[2], fixed = TRUE))
})

test_that("page_breaks", {

  temp <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_data(sheet = 1, x = iris)

  wb$add_page_break(sheet = 1, row = 10)
  wb$add_page_break(sheet = 1, row = 20)
  wb$add_page_break(sheet = 1, col = 2)

  rbrk <- c("<brk id=\"10\" max=\"16383\" man=\"1\"/>",
            "<brk id=\"20\" max=\"16383\" man=\"1\"/>"
  )
  cbrk <- "<brk id=\"2\" max=\"1048575\" man=\"1\"/>"

  expect_equal(rbrk, wb$worksheets[[1]]$rowBreaks)
  expect_equal(cbrk, wb$worksheets[[1]]$colBreaks)

  wb$save(temp)
  wb2 <- wb_load(temp)

  expect_equal(
    wb$worksheets[[1]]$colBreaks,
    wb2$worksheets[[1]]$colBreaks
  )

  expect_equal(
    wb$worksheets[[1]]$rowBreaks,
    wb2$worksheets[[1]]$rowBreaks
  )

})
