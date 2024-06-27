test_that("reading xlsb works", {

  skip_if_offline()

  xlsxFile <- testfile_path("openxlsx2_example.xlsb")

  wb <- wb_load(xlsxFile)
  df_xlsb <- wb_to_df(wb)

  xlsx <- system.file("extdata", "openxlsx2_example.xlsx",
                      package = "openxlsx2")
  df_xlsx <- wb_to_df(xlsx)

  expect_equal(df_xlsb, df_xlsx)

  df_xlsb <- wb_to_df(wb, show_formula = TRUE)
  df_xlsx <- wb_to_df(xlsx, show_formula = TRUE)
  expect_equal(df_xlsb, df_xlsx)

})

test_that("reading complex xlsb works", {

  skip_if_offline()

  xlsxFile <- testfile_path("hyperlink.xlsb")

  expect_message( # larger workbook
    capture_output( # unhandled conditions
      wb <- wb_load(xlsxFile)
    ),
    "importing larger workbook. please wait a moment"
  )

  # chartsheets
  exp <- c(TRUE, FALSE, FALSE, FALSE)
  got <- wb$is_chartsheet
  expect_equal(exp, got)

  # named regions
  exp <- c("Numbers", "characters", "_xlnm._FilterDatabase")
  got <- unique(wb$get_named_regions()$name)
  expect_equal(exp, got)

  # tables
  exp <- "Table1"
  got <- wb$get_tables(sheet = 2)$tab_name
  expect_equal(exp, got)

  # comments
  exp <- c(
    "<t>Jan Marvin Garbuszus:</t>",
    "<t xml:space=\"preserve\">\n</t>",
    "<t>A new note!</t>"
  )
  got <- wb$comments[[1]][[1]]$comment
  expect_equal(exp, got)

  # hyperlinks
  exp <- "Sheet1!E3"
  got <- wb$worksheets[[2]]$hyperlinks[[1]]$location
  expect_equal(exp, got)

})

test_that("worksheets with real world formulas", {

  skip_if_offline()

  xlsxFile <- testfile_path("nhs-core-standards-for-eprr-v6.1.xlsb")

  expect_message( # larger workbook
    capture_output( # unhandled conditions
      suppressWarnings(wb <- wb_load(xlsxFile))
    ),
    "importing larger workbook. please wait a moment"
  )

  exp <- c("Control", "EPRR Core Standards", "Deep dive",
           "Interoperable capabilities ", "Lookups", "Calculations")
  got <- wb$get_sheet_names() %>% names()
  expect_equal(exp, got)

  xlsxFile <- testfile_path("array_fml.xlsb")

  wb <- wb_load(xlsxFile)

  exp <- structure(
    list(
      A = c("1", "2", NA, "6", "#VALUE!"),
      B = c("2", "3", NA, "a", NA),
      C = c(3, 4, NA, NA, NA),
      D = c(4, 5, NA, NA, NA),
      E = c(5, 6, NA, NA, NA)
    ),
    row.names = c(NA, 5L),
    class = "data.frame"
  )
  got <- wb_to_df(wb, col_names = FALSE)
  expect_equal(exp, got)

  exp <- structure(
    list(
      A = c("1", "A1+1", NA, "SUM({\"1\",\"2\",\"3\"})", "SUMIFS(A1:A2,'[1]foo'!B2:B3,{\"A\"})"),
      B = c("A1+1", "B1+1", NA, "{\"a\"}", NA),
      C = c("B1+1", "C1+1", NA, NA, NA),
      D = c("C1+1", "D1+1", NA, NA, NA),
      E = c("D1+1", "E1+1", NA, NA, NA)
    ),
    row.names = c(NA, 5L),
    class = "data.frame"
  )
  got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE)
  expect_equal(exp, got)


  xlsxFile <- testfile_path("match_escape.xlsb")

  wb <- wb_load(xlsxFile)

  exp <- "IF(ISNA(MATCH(G681,{\"Annual\",\"\"\"5\"\"\",\"\"\"1+4\"\"\"},0)),1,MATCH(G681,{\"Annual\",\"\"\"5\"\"\",\"\"\"1+4\"\"\"},0))"
  got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE)$A
  expect_equal(exp, got)

})

test_that("xlsb formulas", {

  fl <- testfile_path("formula_checks.xlsb")
  wb <- wb_load(fl)

  exp <- c("", "D2:E2", "A1,B1", "A1 A2", "1+1", "1-1", "1*1", "1/1",
          "1%", "1^1", "1=1", "1&gt;1", "1&gt;=1", "1&lt;1", "1&lt;=1",
          "1&lt;&gt;1", "+A3", "-R2", "(1)", "SUM(1, )", "1", "2.500000",
          "\"a\"", "\"A\"&amp;\"B\"", "Sheet2!B2", "'[1]Sheet3'!A2")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$f)
  expect_equal(exp, got)

  fl <- testfile_path("formula_checks.xlsx")
  xl <- wb_load(fl)

  expect_equal(
    xl$worksheets[[2]]$sheet_data$cc$f,
    wb$worksheets[[2]]$sheet_data$cc$f
  )

})

test_that("shared formulas are detected correctly", {

  xlsb <- testfile_path("formula_checks.xlsb")
  xlsx <- testfile_path("formula_checks.xlsx")

  wbb <- wb_load(xlsb)
  wbx <- wb_load(xlsx)

  expect_equal(
    wbb$worksheets[[4]]$sheet_data$cc,
    wbx$worksheets[[4]]$sheet_data$cc
  )

})
