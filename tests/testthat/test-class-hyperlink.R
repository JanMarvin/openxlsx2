testsetup()

test_that("encode Hyperlink works", {

  formula_old <- '=HYPERLINK("#Tab_1!" &amp; CELL("address", INDEX(C1:F1, MATCH(A1, C1:F1, 0))), "Go to the selected column")'
  formula_new <- '=HYPERLINK("#Tab_1!" & CELL("address", INDEX(C1:F1, MATCH(A1, C1:F1, 0))), "Go to the selected column")'

  wb <- wb_workbook()$
    add_worksheet("Tab_1", zoom = 80, gridLines = FALSE)$
    add_data(x = rbind(2016:2019), dims = "C1:F1", colNames = FALSE)$
    add_data(x = 2017, dims = "A1", colNames = FALSE)$
    add_data_validation(dims = "A1", type = "list", value = '"2016,2017,2018,2019"')$
    add_formula(dims = "B1", x = formula_old)$
    add_formula(dims = "B2", x = formula_new)

  got <- wb$worksheets[[1]]$sheet_data$cc["11", "f"]
  expect_equal(formula_old, got)

  got <- wb$worksheets[[1]]$sheet_data$cc["12", "f"]
  expect_equal(formula_old, got)

  expect_equal(formula_new, wb_to_df(wb, colNames = FALSE, showFormula = TRUE)[1, "B"])

})

test_that("formulas with hyperlinks works", {

  wb <- wb_workbook()$
    add_worksheet("Tab_1", zoom = 80, gridLines = FALSE)$
    add_data(dims = "C1:F1", x = rbind(2016:2019), colNames = FALSE)$
    add_data(x = 2017, startCol = 1, startRow = 1, colNames = FALSE)$
    add_data_validation(dims = "A1", type = "list", value = '"2016,2017,2018,2019"')$
    add_formula(dims = "B1", x = '=HYPERLINK("#Tab_1!" & CELL("address", INDEX(C1:F1, MATCH(A1, C1:F1, 0))), "Go to the selected column")')$
    add_formula(dims = "B2", x = '=IF(2017 = VALUE(A1), HYPERLINK("github.com","github.com"), A1)')

  exp <- "=IF(2017 = VALUE(A1), HYPERLINK(\"github.com\",\"github.com\"), A1)"
  got <- wb$worksheets[[1]]$sheet_data$cc["12", "f"]
  expect_equal(exp, got)

})


test_that("hyperlinks work", {

  df <- data.frame(
    "Cash" = 1:20, "Cash2" = 31:50,
    "hLink" = "https://r-project.org",
    "hLink2" = LETTERS[1:20],
    "hLink3" = paste0("https://wikipedia.org/wiki/", LETTERS[1:20]),
    "Percentage" = seq(0, 1, length.out = 20)
  )

  # custom target for letters to their wikipedia page
  target <- list(hLink2 = paste0("https://wikipedia.org/wiki/", LETTERS[1:20]))

  wb <- wb_workbook()
  wb$add_worksheet()$add_hyperlink(x = df, target = target, cols = "hLink2")

  class(df$hLink) <- c(class(df$hLink), "hyperlink")
  expect_warning(wb$add_worksheet()$add_hyperlink(x = df, target = target))

  class(df$hLink) <- c("character")
  wb$add_worksheet()$add_hyperlink(x = df, target = target, cols = c("hLink", "hLink2"), as_table = TRUE)

  wb$add_worksheet()$add_hyperlink(x = "noreply@openxlsx2", target = "mailto:noreply@openxlsx2.com", tooltip = "An Invalid E-Mail Adress")

  wb$add_worksheet()$add_hyperlink(x = c("https://r-project.org", "https://cran.r-project.org"), dims = "B2:C2",
                                   tooltip = c("The R-Project Homepage", "CRAN"))

  exp <- c("<hyperlink ref=\"B2\" r:id=\"rId1\" tooltip=\"The R-Project Homepage\"/>",
           "<hyperlink ref=\"C2\" r:id=\"rId3\" tooltip=\"CRAN\"/>")
  got <- wb$worksheets[[5]]$hyperlinks
  expect_equal(exp, got)

  wb$add_worksheet()$add_hyperlink(x = "hyperlinks.xlsx", target = "/Users/janmarvingarbuszus/Source/openxlsx-data/hyperlinks.xlsx", dims = "B2",
                                   tooltip = "my testfile")

  exp <- "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"/Users/janmarvingarbuszus/Source/openxlsx-data/hyperlinks.xlsx\" TargetMode=\"External\"/>"
  got <- wb$worksheets_rels[[6]]
  expect_equal(exp, got)

  wb$add_worksheet()$add_hyperlink(x = "'Sheet 1'!C5", target = "'Sheet 1'!C5", dims = "B2",
                                   tooltip = "An internal reference", is_external = FALSE)$
    add_hyperlink(x = "'Sheet 1'!C5", dims = "B4",
                  tooltip = "An internal reference", is_external = FALSE)

})
