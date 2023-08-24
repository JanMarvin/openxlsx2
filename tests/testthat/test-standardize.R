test_that("standardize works", {

  color <- NULL
  standardize_color_names(colour = "green")
  expect_equal(get("color"), "green")

  tabColor <- NULL
  standardize_color_names(tabColour = "green")
  expect_equal(get("tabColor"), "green")

  camelCase <- NULL
  camel_case <- NULL
  standardize_case_names(camelCase = "green")
  expect_equal(get("camel_case"), "green")

  tab_color <- NULL
  standardize(tabColour = "green")
  expect_equal(get("tab_color"), "green")

})

test_that("deprecation warning works", {

  opt_deprecation <- getOption("openxlsx2.soon_deprecated")

  xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
  wb1 <- wb_load(xlsxFile)

  options("openxlsx2.soon_deprecated" = TRUE)

  expect_warning(
    wb_to_df(wb1, colNames = TRUE),
    "Found camelCase arguments in code. These will be deprecated in the next major release. Consider using: col_names"
  )
  # Do not alter the global state.
  options("openxlsx2.soon_deprecated" = opt_deprecation)

})
