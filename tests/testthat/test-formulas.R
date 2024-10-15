test_that("load various formulas", {

  fl <- testfile_path("formula.xlsx")
  wb <- wb_load(fl)

  expect_false(is.null(wb$metadata))

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
    f_t = "array", f_ref = "A1", f_ca = "", f_si = "", is = "",
    typ = "14")
  got <- wb$worksheets[[1]]$sheet_data$cc[1, ]
  expect_equal(exp, got)

  expect_false(is.null(wb$metadata))

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

test_that("formual escaping works", {

  df_tmp <- data.frame(
    f = "'A&B'!A1",
    g = "'A&amp;B'!A1"
  )
  class(df_tmp$f) <- c(class(df_tmp$f), "formula")
  class(df_tmp$g) <- c(class(df_tmp$g), "formula")

  wb <- wb_workbook()$
    add_worksheet("A&B")$
    add_worksheet("Fml")$
    add_data(x = df_tmp, col_names = FALSE)$
    add_formula(dims = "A2", x = "'A&B'!A1")$
    add_formula(dims = "A3", x = "SUM('A&B'!A1)", array = TRUE)

  expect_warning(wb$add_formula(dims = "A4", x = "SUM('A&B'!A1)", cm = TRUE))

  exp <- c("'A&amp;B'!A1", "'A&amp;B'!A1", "'A&amp;B'!A1", "SUM('A&amp;B'!A1)", "SUM('A&amp;B'!A1)")
  got <- wb$worksheets[[2]]$sheet_data$cc$f
  expect_equal(exp, got)

})

test_that("writing array vectors works", {
  fml <- function(x) {
    paste0("=", x, "+", x)
  }

  wb <- wb_workbook()

  ### write vector of array formulas
  wb$add_worksheet()$add_formula(x = fml(seq_len(10)), array = TRUE, dims = "B2")

  ## check
  cc <- wb$worksheets[[1]]$sheet_data$cc

  exp <- c("B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11")
  got <- cc[cc$f_t == "array", c("f_ref")]
  expect_equal(exp, got)

  ### write array formula as data frame class
  df <- data.frame(
    formula       = fml(seq_len(10)),
    array_formula = fml(seq_len(10)),
    character     = fml(seq_len(10))
  )

  class(df$formula) <- c("formula", class(df$formula))
  class(df$array_formula) <- c("array_formula", class(df$array_formula))

  wb$add_worksheet()$add_data(x = df, dims = "B2")

  ## check
  cc <- wb$worksheets[[2]]$sheet_data$cc

  exp <- c("C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12")
  got <- cc[cc$f_t == "array", c("f_ref")]
  expect_equal(exp, got)

})

test_that("array formula detection works", {
  wb <- wb_workbook()$add_worksheet()$
    add_formula(dims = "A1", x = "{1+1}")
  cc <- wb$worksheets[[1]]$sheet_data$cc

  exp <- "A1"
  got <- cc[cc$f_t == "array", "f_ref"]
  expect_equal(exp, got)
})

test_that("writing shared formulas works", {
  df <- data.frame(
    x = 1:5,
    y = 1:5 * 2
  )

  wb <-  wb_workbook()$add_worksheet()$add_data(x = df)

  wb$add_formula(
    x      = "=A2/B2",
    dims   = "C2:C6",
    array  = FALSE,
    shared = TRUE
  )

  cc <- wb$worksheets[[1]]$sheet_data$cc
  cc <- cc[cc$c_r == "C", ]

  exp <- c("=A2/B2", "", "", "", "")
  got <- cc$f
  expect_equal(exp, got)

  exp <- c("shared")
  got <- unique(cc$f_t)
  expect_equal(exp, got)

  wb$add_formula(
    x      = "=A$2/B$2",
    dims   = "D2:D6",
    array  = FALSE,
    shared = TRUE
  )

  cc <- wb$worksheets[[1]]$sheet_data$cc
  cc <- cc[cc$c_r == "D", ]

  exp <- c("=A$2/B$2", "", "", "", "")
  got <- cc$f
  expect_equal(exp, got)

  exp <- c("1")
  got <- unique(cc$f_si)
  expect_equal(exp, got)

})

test_that("increase formula dims if required", {

  fml <- c("SUM(A2:B2)", "SUM(A3:B3)")

  # This only handles single cells, if C2 is passed and length(x) > 1
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = matrix(1:4, 2, 2))

  wb1 <- wb_add_formula(wb, dims = "C2", x = fml)
  wb2 <- wb_add_formula(wb, dims = "C2:D2", x = fml)

  expect_equal(
    wb1$worksheets[[1]]$sheet_data$cc,
    wb2$worksheets[[1]]$sheet_data$cc
  )

})

test_that("registering formulas works", {

  fml <- "_xlfn.LAMBDA(TODAY() - 1)"
  wb <- wb_workbook()$add_worksheet()

  expect_message(wb$add_formula(x = c(YESTERDAY = fml)), "formula registered to the workbook")
  expect_error(wb$add_formula(x = c(YESTERDAY = fml)), "named regions cannot be duplicates")
  expect_equal(wb$get_named_regions()$value, fml)

  wb <- wb_add_formula(wb, x = "YESTERDAY()", name = "YSTRDY", array = TRUE)
  expect_equal(wb$get_named_regions()$name, c("YESTERDAY", "YSTRDY"))

})
