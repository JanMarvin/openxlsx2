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
  exp <- "<t>Jan Marvin Garbuszus:\nA new note!</t>"
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

})
