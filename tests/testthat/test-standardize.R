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

  xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
  wb1 <- wb_load(xlsxFile)

  op <- options("openxlsx2.soon_deprecated" = TRUE)
  on.exit(options(op), add = TRUE)

  expect_warning(
    wb_to_df(wb1, colNames = TRUE),
    "Found camelCase arguments in code. These will be deprecated in the next major release. Consider using: col_names"
  )

})

test_that("functions consuming ... warn on unknown arguments (#1646)", {

  wb <- wb_workbook()$add_worksheet()$add_data(x = mtcars[1:3, ])

  # these used to absorb unknown arguments silently (no warning at all).
  # NB: use names that are not prefixes of a real argument, otherwise R's
  # partial matching would bind them before they reach `...`.
  expect_warning(
    wb$add_conditional_formatting(dims = "A2:A4", rule = ">3", nonexistent = 1),
    "unused arguments \\(nonexistent\\)"
  )
  expect_warning(
    wb$merge_cells(dims = "A1:B1", directon = "row"),
    "unused arguments \\(directon\\)"
  )
  expect_warning(
    wb$unmerge_cells(dims = "A1:B1", nonexistent = 1),
    "unused arguments \\(nonexistent\\)"
  )
  expect_warning(
    wb$set_base_colors(nonexistent = 1),
    "unused arguments \\(nonexistent\\)"
  )

  # a valid argument is not flagged
  expect_no_warning(wb$merge_cells(dims = "F2:G2"))
})
