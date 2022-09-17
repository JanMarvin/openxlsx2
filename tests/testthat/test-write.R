test_that("write_formula", {

  set.seed(123)
  df <- data.frame(C = rnorm(10), D = rnorm(10))

  # array formula for a single cell
  exp <- structure(
    list(r = "E2", row_r = "2", c_r = "E", c_s = "",
         c_t = "", c_cm = "",
         c_ph = "", c_vm = "",
         v = "", f = "SUM(C2:C11*D2:D11)",
         f_t = "array", f_ref = "E2:E2",
         f_ca = "", f_si = "",
         is = "", typ = ""),
    row.names = 23L, class = "data.frame")

  # write data add array formula later
  wb <- wb_workbook()
  wb <- wb_add_worksheet(wb, "df")
  wb$add_data("df", df, startCol = "C")
  write_formula(wb, "df", startCol = "E", startRow = "2",
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)

  cc <- wb$worksheets[[1]]$sheet_data$cc
  got <- cc[cc$row_r == "2" & cc$c_r == "E",]
  expect_equal(exp[1:16], got[1:16])


  rownames(exp) <- 1L
  # write formula first add data later
  wb <- wb_workbook()
  wb <- wb_add_worksheet(wb, "df")
  write_formula(wb, "df", startCol = "E", startRow = "2",
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)
  wb$add_data("df", df, startCol = "C")

  cc <- wb$worksheets[[1]]$sheet_data$cc
  got <- cc[cc$row_r == "2" & cc$c_r == "E",]
  expect_equal(exp[1:11], got[1:11])

})

test_that("silent with numfmt option", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")
  wb$add_worksheet("S2")

  wb$add_data_table("S1", x = iris)
  expect_warning(
    wb$add_data_table("S2",
                 x = mtcars, xy = c("B", 3), rowNames = TRUE,
                 tableStyle = "TableStyleLight9")
  )

  # [1:4] to ignore factor
  expect_equal(iris[1:4], wb_to_df(wb, "S1")[1:4], ignore_attr = TRUE)
  expect_equal(iris[1:4], wb_to_df(wb, "S1")[1:4], ignore_attr = TRUE)

  # handle rownames
  got <- wb_to_df(wb, "S2", rowNames = TRUE)
  attr(got, "tt") <- NULL
  attr(got, "types") <- NULL
  expect_equal(mtcars, got)
  expect_equal(rownames(mtcars), rownames(got))

})

test_that("test options", {

  ops <- options()
  tmp <- temp_xlsx()
  wb_workbook()$add_worksheet("Sheet 1")$add_data("Sheet 1", cars)
  ops2 <- options()

  # adding data to the worksheet should not alter the global options
  expect_equal(ops, ops2)

})

test_that("missing x is caught early [246]", {
  expect_error(
    wb_workbook()$add_data(mtcars),
    "`x` is missing"
  )
})

test_that("update_cells", {

  ## exactly the same
  data <- mtcars
  wb <- wb_workbook()$add_worksheet()$add_data(x = data)
  cc1 <- wb$worksheets[[1]]$sheet_data$cc

  wb$add_data(x = data)
  cc2 <- wb$worksheets[[1]]$sheet_data$cc

  all.equal(cc1, cc2)

  ## write na.strings
  data <- matrix(NA, 2, 2)
  wb <- wb_workbook()$add_worksheet()$add_data(x = data)$add_data(x = data, na.strings = "N/A")

  exp <- c("<is><t>V1</t></is>", "<is><t>V2</t></is>", "<is><t>N/A</t></is>")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$is)
  expect_equal(exp, got)

  ### write logical
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  wb1 <- wb_load(xlsxFile)

  data <- head(wb_to_df(wb1, sheet = 3))
  wb <- wb_workbook()$add_worksheet()$add_data(x = data)$add_data(x = data)

  exp <- c("inlineStr", "", "b", "e")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$c_t)
  expect_equal(exp, got)


  set.seed(123)
  df <- data.frame(C = rnorm(10), D = rnorm(10))

  wb <- wb_workbook()$
    add_worksheet("df")$
    add_data(x = df, startCol = "C")
  # TODO add_formula()
  write_formula(wb, "df", startCol = "E", startRow = "2",
                x = "SUM(C2:C11*D2:D11)",
                array = TRUE)
  write_formula(wb, "df", x = "C3 + D3", startCol = "E", startRow = 3)
  x <- c(google = "https://www.google.com")
  class(x) <- "hyperlink"
  wb$add_data(sheet = "df", x = x, startCol = "E", startRow = 4)


  exp <- structure(
    list(c_t = c("", "str", ""),
         f = c("SUM(C2:C11*D2:D11)", "C3 + D3", "=HYPERLINK(\"https://www.google.com\")"),
         f_t = c("array", "", "")),
    row.names = c("23", "110", "111"), class = "data.frame")
  got <- wb$worksheets[[1]]$sheet_data$cc[c(5,8,11), c("c_t", "f", "f_t")]
  expect_equal(exp, got)

})

test_that("write dims", {

  # create a workbook
  wb <- wb_workbook()$
    add_worksheet()$add_data(dims = "B2:C3", x = matrix(1:4, 2, 2), colNames = FALSE)$
    add_worksheet()$add_data_table(dims = "B:C", x = as.data.frame(matrix(1:4, 2, 2)))$
    add_worksheet()$add_formula(dims = "B3", x = "42")

  s1 <- wb_to_df(wb, 1, colNames = FALSE)
  s2 <- wb_to_df(wb, 2, colNames = FALSE)
  s3 <- wb_to_df(wb, 3, colNames = FALSE)

  expect_equal(rownames(s1), c("2", "3"))
  expect_equal(rownames(s2), c("1", "2", "3"))
  expect_equal(rownames(s3), c("3"))

  expect_equal(colnames(s1), c("B", "C"))
  expect_equal(colnames(s2), c("B", "C"))
  expect_equal(colnames(s3), c("B"))

})

test_that("write data.table class", {

  df <- mtcars
  class(df) <- c("data.table", "data.frame")

  tmp <- temp_xlsx()
  expect_silent(write_xlsx(df, tmp))
  expect_equal(mtcars, read_xlsx(tmp), ignore_attr = TRUE)

})
