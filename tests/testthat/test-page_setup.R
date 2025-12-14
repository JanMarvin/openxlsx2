test_that("Page setup", {
  wb <- wb_workbook()
  wb$add_worksheet("s1")
  wb$add_worksheet("s2")
  wb$add_worksheet("s3")

  wb$page_setup(
    sheet = "s1", orientation = "landscape", scale = 100, left = 0.1,
    right = 0.1, top = 0.75, bottom = 0.75, header = 0.1, footer = 0.1,
    fit_to_width = TRUE, fit_to_height = TRUE, paper_size = 1,
    summary_row = "below", summary_col = "right"
  )

  wb$page_setup(
    sheet = 2, orientation = "landscape", scale = 100, left = 0.1,
    right = 0.1, top = 0.75, bottom = 0.75, header = 0.1, footer = 0.1,
    fit_to_width = TRUE, fit_to_height = TRUE, paper_size = 1,
    summary_row = "below", summary_col = "right"
  )

  wb$set_page_setup(
    sheet = "s3", tab_color = wb_color("green")
  )

  expect_equal(wb$worksheets[[1]]$pageSetup, wb$worksheets[[2]]$pageSetup)

  v <- gsub(" ", "", wb$worksheets[[1]]$pageSetup, fixed = TRUE)
  expect_match(v, 'paperSize="1"')
  expect_match(v, 'orientation="landscape"')
  expect_match(v, 'fitToWidth="1"')
  expect_match(v, 'fitToHeight="1"')

  pr <- wb$worksheets[[1]]$sheetPr
  expect_match(pr, '<outlinePr summaryBelow="1" summaryRight="1"/>', fixed = TRUE)

  pr <- wb$worksheets[[3]]$sheetPr
  expect_equal("<sheetPr><tabColor rgb=\"FF00FF00\"/></sheetPr>", pr)

  wb$set_page_setup(
    sheet = "s3", tab_colour = ""
  )

  pr <- wb$worksheets[[3]]$sheetPr
  expect_equal(pr, "<sheetPr/>")

})

test_that("page_breaks", {

  temp <- temp_xlsx()
  on.exit(unlink(temp), add = TRUE)

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

  expect_equal(wb$worksheets[[1]]$rowBreaks, rbrk)
  expect_equal(wb$worksheets[[1]]$colBreaks, cbrk)

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
