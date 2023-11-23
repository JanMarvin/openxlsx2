testsetup()

test_that("class Hyperlink works", {
  expect_null(assert_hyperlink(wb_hyperlink()))
})

test_that("encode Hyperlink works", {

  formula_old <- '=HYPERLINK("#Tab_1!" &amp; CELL("address", INDEX(C1:F1, MATCH(A1, C1:F1, 0))), "Go to the selected column")'
  formula_new <- '=HYPERLINK("#Tab_1!" & CELL("address", INDEX(C1:F1, MATCH(A1, C1:F1, 0))), "Go to the selected column")'

  wb <- wb_workbook()$
    add_worksheet("Tab_1", zoom = 80, gridLines = FALSE)$
    add_data(x = rbind(2016:2019), dims = "C1:F1", colNames = FALSE)$
    add_data(x = 2017, dims = "A1", colNames = FALSE)$
    add_data_validation(dims = "A1", type = "list", value = '"2016,2017,2018,2019"')$
    add_formula(dims = "B1", x = formula_old)$
    add_formula(dims = "B2", x = formula_new)

  got <- wb$worksheets[[1]]$sheet_data$cc["11", "f"]
  expect_equal(formula_old, got)

  got <- wb$worksheets[[1]]$sheet_data$cc["12", "f"]
  expect_equal(formula_old, got)

  expect_equal(formula_new, wb_to_df(wb, colNames = FALSE, showFormula = TRUE)[1, "B"])

})

test_that("formulas with hyperlinks works", {

  wb <- wb_workbook()$
    add_worksheet("Tab_1", zoom = 80, gridLines = FALSE)$
    add_data(dims = "C1:F1", x = rbind(2016:2019), colNames = FALSE)$
    add_data(x = 2017, startCol = 1, startRow = 1, colNames = FALSE)$
    add_data_validation(dims = "A1", type = "list", value = '"2016,2017,2018,2019"')$
    add_formula(dims = "B1", x = '=HYPERLINK("#Tab_1!" & CELL("address", INDEX(C1:F1, MATCH(A1, C1:F1, 0))), "Go to the selected column")')$
    add_formula(dims = "B2", x = '=IF(2017 = VALUE(A1), HYPERLINK("github.com","github.com"), A1)')

  exp <- "=IF(2017 = VALUE(A1), HYPERLINK(\"github.com\",\"github.com\"), A1)"
  got <- wb$worksheets[[1]]$sheet_data$cc["12", "f"]
  expect_equal(exp, got)

})
