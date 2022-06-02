test_that("write_formula", {

  set.seed(123)
  df <- data.frame(C = rnorm(10), D = rnorm(10))

  # array formula for a single cell
  exp <- structure(
    list(row_r = "2", c_r = "E", c_s = "",
         c_t = "", c_cm = "",
         c_ph = "", c_vm = "",
         v = "", f = "SUM(C2:C11*D2:D11)",
         f_t = "array", f_ref = "E2:E2",
         f_ca = "", f_si = "",
         is = "", typ = NA_character_, r = "E2"),
    row.names = "3", class = "data.frame")

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


  rownames(exp) <- "31"
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
  wb$add_data_table("S2",
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


test_that("test options", {

  ops <- options()
  tmp <- temp_xlsx()
  wb_workbook()$add_worksheet("Sheet 1")$add_data("Sheet 1", cars)
  ops2 <- options()

  # adding data to the worksheet should not alter the global options
  expect_equal(ops, ops2)

})
