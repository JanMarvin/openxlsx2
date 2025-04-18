test_that("load various formulas", {

  fl <- testfile_path("formula.xlsx")
  wb <- wb_load(fl)

  expect_false(is.null(wb$metadata))

  cc <- wb$worksheets[[1]]$sheet_data$cc

  expect_identical(cc[1, "c_cm"], "1")
  expect_identical(cc[1, "f_attr"], "t=\"array\" ref=\"A1\" ") # FIXME: why the trailing whitespace?

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
    c_cm = "1", v = "", f = "SUM(ABS(A2:A11))",
    f_attr = "t=\"array\" ref=\"A1\"", is = "",
    stringsAsFactors = FALSE)
  got <- wb$worksheets[[1]]$sheet_data$cc[1, ]
  expect_equal(exp, got)

  expect_false(is.null(wb$metadata))

})

test_that("setting ref works", {

  m1 <- matrix(1:6, ncol = 2)
  m2 <- matrix(7:12, nrow = 2)

  wb <- wb_workbook()$add_worksheet()$
    add_data(x = m1, start_col = 1)$
    add_data(x = m2, start_col = 4)$
    add_formula(dims = "H1:J3", x = "MMULT(A2:B4, D2:F3)", array = TRUE)

  cc <- wb$worksheets[[1]]$sheet_data$cc

  exp <- "t=\"array\" ref=\"H1:J3\""
  got <- cc[cc$r == "H1", "f_attr"]
  expect_equal(exp, got)

})

test_that("formual escaping works", {

  df_tmp <- data.frame(
    f = "'A&B'!A1",
    g = "'A&amp;B'!A1",
    stringsAsFactors = FALSE
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

  expect_silent(wb$save(temp_xlsx()))

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
  ff <- rbindlist(xml_attr(paste0("<f ", cc$f_attr, "/>"), "f"))

  exp <- sprintf("t=\"array\" ref=\"%s\"", c("B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11"))
  got <- cc[ff$t == "array", c("f_attr")]
  expect_equal(exp, got)

  ### write array formula as data frame class
  df <- data.frame(
    formula       = fml(seq_len(10)),
    array_formula = fml(seq_len(10)),
    character     = fml(seq_len(10)),
    stringsAsFactors = FALSE
  )

  class(df$formula) <- c("formula", class(df$formula))
  class(df$array_formula) <- c("array_formula", class(df$array_formula))

  wb$add_worksheet()$add_data(x = df, dims = "B2")

  ## check
  cc <- wb$worksheets[[2]]$sheet_data$cc
  ff <- rbindlist(xml_attr(paste0("<f ", cc$f_attr, "/>"), "f"))

  exp <- sprintf("t=\"array\" ref=\"%s\"", c("C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12"))
  got <- cc[ff$t == "array", c("f_attr")]
  expect_equal(exp, got)

})

test_that("array formula detection works", {
  wb <- wb_workbook()$add_worksheet()$
    add_formula(dims = "A1", x = "{1+1}")
  cc <- wb$worksheets[[1]]$sheet_data$cc
  ff <- rbindlist(xml_attr(paste0("<f ", cc$f_attr, "/>"), "f"))

  exp <- sprintf("t=\"array\" ref=\"%s\"", "A1")
  got <- cc[ff$t == "array", "f_attr"]
  expect_equal(exp, got)
})

test_that("writing shared formulas works", {
  # somehow this test is failing in Rstudio and I do not know why
  df <- data.frame(
    x = 1:5,
    y = 1:5 * 2,
    stringsAsFactors = FALSE
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

  exp <- c("t=\"shared\" ref=\"C2:C6\" si=\"0\"", "t=\"shared\" si=\"0\"")
  got <- unique(cc$f_attr)
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

  exp <- c("t=\"shared\" ref=\"D2:D6\" si=\"1\"", "t=\"shared\" si=\"1\"")
  got <- unique(cc$f_attr)
  expect_equal(exp, got)

  wb <- wb_workbook()$add_worksheet()
  wb$add_formula(x = "1", dims = "A1:B1", shared = TRUE)

  exp <- c("t=\"shared\" ref=\"A1:B1\" si=\"0\"", "t=\"shared\" si=\"0\"")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$f_attr)
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

test_that("writing cm formulas and writing to sheets with cm formulas works", {

  wb <- wb_workbook()$add_worksheet()
  wb$add_data(x = c(1, 1), dims = "D1:D2")
  wb$add_data(x = c(1, 1), dims = "D4:E4")
  expect_warning(wb$add_formula(dims = "G1", x = "MMULT(D4:E4,D1:D2)", cm = TRUE))

  expect_equal("1", wb$worksheets[[1]]$sheet_data$cc$c_cm[2])

  expect_silent(wb$add_data(x = 1))

})

test_that("formula with names work as documented", {

  v1 <- rep("https://CRAN.R-project.org/", 4)
  names(v1) <- paste0("Hyperlink", 1:4) # Optional: names will be used as display text
  class(v1) <- "hyperlink"

  v2 <- rep("https://CRAN.R-project.org/", 4)
  class(v2) <- "hyperlink"

  wb <- wb_workbook()$add_worksheet()
  wb$add_data(x = v1, dims = "A1")
  wb$add_data(x = v2, dims = "B1")


  exp <- data.frame(
    A = c(
      "=HYPERLINK(\"https://CRAN.R-project.org/\", \"Hyperlink1\")",
      "=HYPERLINK(\"https://CRAN.R-project.org/\", \"Hyperlink2\")",
      "=HYPERLINK(\"https://CRAN.R-project.org/\", \"Hyperlink3\")",
      "=HYPERLINK(\"https://CRAN.R-project.org/\", \"Hyperlink4\")"
    ),
    B = c(
      "=HYPERLINK(\"https://CRAN.R-project.org/\")",
      "=HYPERLINK(\"https://CRAN.R-project.org/\")",
      "=HYPERLINK(\"https://CRAN.R-project.org/\")",
      "=HYPERLINK(\"https://CRAN.R-project.org/\")"
    )
  )
  class(exp$A) <- c("character", "formula")
  class(exp$B) <- c("character", "formula")
  got <- wb$to_df(show_formula = TRUE, col_names = FALSE)
  expect_equal(exp, got)

})

test_that("normalizing spreadsheet formulas works", {
  fml <- c(
    'SUM(A1;B1; C1)',
    'SUM(A1,B1,C1)',
    '"Hello;
  World" & 1',
    'TEXT(TODAY(),"MM/DD/YY")',
    'IF((SUM(A1,B1) & "a;b")<>"",(SUM(A1,B1) & "a;b");"")',
    '',
    '  SUM(A1;B1;C1)  '
  )

  wb <- wb_workbook()$add_worksheet()$
    add_data(x = matrix(1:9, 3, 3), col_names = FALSE)$
    add_formula(x = fml, dims = "D2")

  exp <- structure(
    c("SUM(A1,B1, C1)",
      "SUM(A1,B1,C1)",
      "\"Hello;\n  World\" & 1",
      "TEXT(TODAY(),\"MM/DD/YY\")",
      "IF((SUM(A1,B1) & \"a;b\")<>\"\",(SUM(A1,B1) & \"a;b\"),\"\")",
      NA,
      "  SUM(A1,B1,C1)  "
    ),
    class = c("character", "formula")
  )
  got <- wb$to_df(show_formula = TRUE, col_names = FALSE, start_row = 2, cols = "D")$D
  expect_equal(exp, got)
})

test_that("cm works more like array", {
  wb <- wb_workbook()$add_worksheet()$
    add_data(x = cars)$
    add_data(dims = "D1", x = "Unique Values of Speed")

  expect_warning(
    wb$add_formula(
      dims = wb_dims(x = unique(cars$speed), from_dims = "D2"),
      x = paste0("_xlfn.UNIQUE(", wb_dims(x = cars, cols = "speed"), ")"),
      cm = TRUE
    ),
    "modifications with cm formulas are experimental"
  )

  got <- table(wb$worksheets[[1]]$sheet_data$cc$f)[["_xlfn.UNIQUE(A2:A51)"]]
  expect_equal(1L, got)
})
