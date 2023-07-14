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
