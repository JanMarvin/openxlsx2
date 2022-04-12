test_that("writeFormula", {

  set.seed(123)
  df <- data.frame(C = rnorm(10), D = rnorm(10))

  # array formula for a single cell
  exp <- structure(
    list(row_r = "2", c_r = "E", c_s = "_openxlsx_NA_",
         c_t = "_openxlsx_NA_", v = "_openxlsx_NA_",
         f = "SUM(C2:C11*D2:D11)",  f_t = "array", f_ref = "E2:E2",
         f_ca = "_openxlsx_NA_", f_si = "_openxlsx_NA_",
         is = "_openxlsx_NA_", typ = NA_character_, r = "E2"),
    row.names = "3", class = "data.frame")

  # write data add array formula later
  wb <- wb_workbook()
  wb <- wb_add_worksheet(wb, "df")
  writeData(wb, "df", df, startCol = "C")
  writeFormula(wb, "df", startCol = "E", startRow = "2",
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)

  cc <- wb$worksheets[[1]]$sheet_data$cc
  got <- cc[cc$row_r == "2" & cc$c_r == "E",]
  expect_equal(exp[1:11], got[1:11])


  rownames(exp) <- "31"
  # write formula first add data later
  wb <- wb_workbook()
  wb <- wb_add_worksheet(wb, "df")
  writeFormula(wb, "df", startCol = "E", startRow = "2",
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)
  writeData(wb, "df", df, startCol = "C")

  cc <- wb$worksheets[[1]]$sheet_data$cc
  got <- cc[cc$row_r == "2" & cc$c_r == "E",]
  expect_equal(exp[1:11], got[1:11])

})

test_that("silent with numfmt option", {

  wb <- wb_workbook()
  wb$addWorksheet("S1")
  wb$addWorksheet("S2")

  writeDataTable(wb, "S1", x = iris)
  writeDataTable(wb, "S2",
                 x = mtcars, xy = c("B", 3), rowNames = TRUE,
                 tableStyle = "TableStyleLight9"
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


test_that("write xlsx", {

  tmp <- temp_xlsx()
  df <- data.frame(a = 1:26, b = letters)

  # expect_silent(write.xlsx(df, tmp, tabColour = "#4F81BD"))
  expect_error(write.xlsx(df, tmp, asTable = "YES"))
  expect_error(write.xlsx(df, tmp, sheetName = paste0(letters, letters, collapse = "")))
  expect_error(write.xlsx(df, tmp, zoom = "FULL"))
  expect_silent(write.xlsx(df, tmp, zoom = 200))
  expect_silent(write.xlsx(x = list("S1" = df, "S2" = df), tmp, sheetName = c("Sheet1", "Sheet2")))
  expect_silent(write.xlsx(x = list("S1" = df, "S2" = df), file = tmp))
  expect_error(write.xlsx(df, tmp, gridLines = "YES"))
  expect_silent(write.xlsx(df, tmp, gridLines = FALSE))
  expect_error(write.xlsx(df, tmp, overwrite = FALSE))
  expect_error(write.xlsx(df, tmp, overwrite = "NO"))
  expect_silent(write.xlsx(df, tmp, withFilter = FALSE))
  expect_silent(write.xlsx(df, tmp, withFilter = TRUE))
  expect_error(write.xlsx(df, tmp, withFilter = "NO"))
  ## FIXME both do not work as expected
  # expect_error(write.xlsx(df, tmp, startRow = "A"))
  # expect_error(write.xlsx(df, tmp, startCol = "2"))
  expect_error(write.xlsx(df, tmp, col.names = "NO"))
  expect_silent(write.xlsx(df, tmp, col.names = TRUE))
  expect_error(write.xlsx(df, tmp, colNames = "NO"))
  expect_silent(write.xlsx(df, tmp, colNames = TRUE))
  expect_error(write.xlsx(df, tmp, row.names = "NO"))
  expect_silent(write.xlsx(df, tmp, row.names = TRUE))
  expect_error(write.xlsx(df, tmp, rowNames = "NO"))
  expect_silent(write.xlsx(df, tmp, rowNames = TRUE))
  expect_error(write.xlsx(df, tmp, xy = "A2"))
  expect_silent(write.xlsx(df, tmp, xy = c(1, 2)))
  expect_silent(write.xlsx(df, tmp, xy = c(1, 2)))
  expect_silent(write.xlsx(df, tmp, colWidth = "auto"))
  # expect_error(write.xlsx(df, tmp, keepNA = "yes"))
  # test works but does not work as intended
  expect_silent(write.xlsx(data.frame(x = NA), tmp, keepNA = TRUE))
  expect_silent(write.xlsx(data.frame(x = NA), tmp, keepNA = TRUE))
  # test works but does not work as intended
  expect_silent(write.xlsx(data.frame(x = NA), tmp, na.string = "N-A"))
  expect_silent(write.xlsx(df, tmp, asTable = TRUE, tableStyle = "TableStyleLight9"))

})
