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

test_that("na.strings option works", {

  op <- options("openxlsx2.na.strings" = fmt_txt("N/A", italic = TRUE, color = wb_color("gray")))
  on.exit(options(op), add = TRUE)

  ## create a data set with missings
  dd <- mtcars
  dd[dd < 3] <- NA

  wb <- write_xlsx(x = dd)

  exp <- "<is><r><rPr><i/><color rgb=\"FFBEBEBE\"/></rPr><t>N/A</t></r></is>"
  got <- wb$worksheets[[1]]$sheet_data$cc[17, "is"]
  expect_equal(got, exp)

  wb <- wb_workbook()
  wb$add_worksheet()
  wb$add_data(x = dd, na.strings = "", row_names = TRUE)
  res <- wb$to_df(row_names = TRUE)
  expect_equal(res, dd)

  wb$add_worksheet()
  wb$add_data(x = dd, row_names = TRUE)
  res2 <- wb$to_df(row_names = TRUE, na.strings = "N/A")
  expect_equal(res2, dd)
})

test_that("na option works", {
  op <- options("openxlsx2.na" = fmt_txt("N/A", italic = TRUE, color = wb_color("gray")))
  on.exit(options(op), add = TRUE)

  ## create a data set with missings
  dd <- mtcars
  dd[dd < 3] <- NA

  wb <- write_xlsx(x = dd)

  exp <- "<is><r><rPr><i/><color rgb=\"FFBEBEBE\"/></rPr><t>N/A</t></r></is>"
  got <- wb$worksheets[[1]]$sheet_data$cc[17, "is"]
  expect_equal(got, exp)

  wb <- wb_workbook()
  wb$add_worksheet()
  wb$add_data(x = dd, na.strings = "", row_names = TRUE)
  res <- wb$to_df(row_names = TRUE)
  expect_equal(res, dd)

  wb$add_worksheet()
  wb$add_data(x = dd, row_names = TRUE)
  res2 <- wb$to_df(row_names = TRUE, na.strings = "N/A")
  expect_equal(res2, dd)

})

test_that("na option works", {
  op <- options("openxlsx2.na" = "_openxlsx_NULL")
  on.exit(options(op), add = TRUE)
  wb <- write_xlsx(matrix(NA, 2, 2))
  expect_equal(unique(wb$worksheets[[1]]$sheet_data$cc$c_t), c("inlineStr", ""))
})
