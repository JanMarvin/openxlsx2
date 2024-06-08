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
    "<t xml:space=\"preserve\">Jan Marvin Garbuszus:</t>",
    "<t xml:space=\"preserve\">\n</t>",
    "<t xml:space=\"preserve\">A new note!</t>"
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

})
