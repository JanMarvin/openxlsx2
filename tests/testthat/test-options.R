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

test_that("na_inf_pos, na_inf_neg and na_nan options replace Inf, -Inf and NaN", {

  x <- data.frame(a = c(1, Inf, -Inf, NaN, NA))

  # default: unchanged behavior
  wb <- wb_workbook()$add_worksheet()$add_data(x = x, inline_strings = FALSE)
  cc <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(cc$c_t[3:5], c("e", "e", "e"))
  expect_equal(cc$v[3:5], c("#NUM!", "#NUM!", "#VALUE!"))

  # strings
  op <- options(
    "openxlsx2.na_inf_pos" = ">5",
    "openxlsx2.na_inf_neg" = "<5",
    "openxlsx2.na_nan"     = "n/a"
  )
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()$add_worksheet()$add_data(x = x)
  got <- wb_to_df(wb)
  expect_equal(got$a, c("1", ">5", "<5", "n/a", NA_character_))

  # shared strings
  wb <- wb_workbook()$add_worksheet()$add_data(x = x, inline_strings = FALSE)
  got <- wb_to_df(wb)
  expect_equal(got$a, c("1", ">5", "<5", "n/a", NA_character_))

  # numbers are written as numbers
  options(
    "openxlsx2.na_inf_pos" = 1e6,
    "openxlsx2.na_inf_neg" = -1e6,
    "openxlsx2.na_nan"     = NULL
  )
  wb <- wb_workbook()$add_worksheet()$add_data(x = x)
  got <- wb_to_df(wb)
  # the column is read as character since it contains error cells
  expect_equal(as.numeric(got$a[1:3]), c(1, 1e6, -1e6))

  # options are independent: only Inf replaced
  options("openxlsx2.na_inf_pos" = "hi", "openxlsx2.na_inf_neg" = NULL)
  wb <- wb_workbook()$add_worksheet()$add_data(x = x, inline_strings = FALSE)
  cc <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(cc$v[4:5], c("#NUM!", "#VALUE!"))

  # character columns containing "Inf" are untouched
  options("openxlsx2.na_inf_pos" = "hi")
  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(a = c("Inf", "-Inf", "NaN")))
  expect_equal(wb_to_df(wb)$a, c("Inf", "-Inf", "NaN"))

  # invalid
  options("openxlsx2.na_inf_pos" = c("a", "b"))
  expect_error(wb_workbook()$add_worksheet()$add_data(x = x), "option `openxlsx2.na_inf_pos`")
  options("openxlsx2.na_inf_pos" = Inf)
  expect_error(wb_workbook()$add_worksheet()$add_data(x = x), "option `openxlsx2.na_inf_pos`")
})
