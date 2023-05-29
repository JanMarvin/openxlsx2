test_that("load various formulas", {

  fl <- testfile_path("formula.xlsx")
  wb <- wb_load(fl)

  expect_true(!is.null(wb$metadata))

  cc <- wb$worksheets[[1]]$sheet_data$cc

  expect_identical(cc[1, "c_cm"], "1")
  expect_identical(cc[1, "f_t"], "array")

  tmp <- temp_xlsx()
  expect_silent(wb$save(tmp))

  wb1 <- wb_load(tmp)

  expect_identical(
    wb$worksheets[[1]]$sheet_data$cc,
    wb1$worksheets[[1]]$sheet_data$cc
  )

  expect_identical(
    wb$metadata,
    wb1$metadata
  )

})

test_that("writing formulas with cell metadata works", {

  wb <- wb_workbook()$
    add_worksheet()

  expect_warning(
    wb$add_formula(x = 'SUM(ABS(A2:A11))', cm = TRUE),
    "modifications with cm formulas are experimental. use at own risk"
  )

  exp <- data.frame(
    r = "A1", row_r = "1", c_r = "A", c_s = "", c_t = "",
    c_cm = "1", c_ph = "", c_vm = "", v = "", f = "SUM(ABS(A2:A11))",
    f_t = "array", f_ref = "A1:A1", f_ca = "", f_si = "", is = "",
    typ = "14")
  got <- wb$worksheets[[1]]$sheet_data$cc[1, ]
  expect_equal(exp, got)

  expect_true(!is.null(wb$metadata))

})

test_that("setting ref works", {

  m1 <- matrix(1:6, ncol = 2)
  m2 <- matrix(7:12, nrow = 2)

  wb <- wb_workbook()$add_worksheet()$
    add_data(x = m1, startCol = 1)$
    add_data(x = m2, startCol = 4)$
    add_formula(dims = "H1:J3", x = "MMULT(A2:B4, D2:F3)", array = TRUE)

  cc <- wb$worksheets[[1]]$sheet_data$cc

  exp <- "H1:J3"
  got <- cc[cc$r == "H1", "f_ref"]
  expect_equal(exp, got)

})
