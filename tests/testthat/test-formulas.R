test_that("load various formulas", {

  fl <- system.file("extdata", "formula.xlsx", package = "openxlsx2")
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
