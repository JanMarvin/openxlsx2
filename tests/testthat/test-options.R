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

test_that("nastrings option works", {

  op <- options("openxlsx2.nastrings" = fmt_txt("N/A", italic = TRUE, color = wb_color("gray")))
  on.exit(options(op), add = TRUE)

  ## create a data set with missings
  dd <- mtcars
  dd[dd < 3] <- NA

  wb <- write_xlsx(x = dd)

  exp <- "<is><r><rPr><i/><color rgb=\"FFBEBEBE\"/></rPr><t>N/A</t></r></is>"
  got <- wb$worksheets[[1]]$sheet_data$cc[17, "is"]
  expect_equal(exp, got)

})
