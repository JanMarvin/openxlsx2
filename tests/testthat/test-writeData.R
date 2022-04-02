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
