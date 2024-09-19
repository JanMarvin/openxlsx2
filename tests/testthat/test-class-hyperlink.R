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

  # custom target for letters to their wikipedia page and a few dims
  target       <- list(hLink2 = paste0("https://wikipedia.org/wiki/", LETTERS[1:20]))
  dims_hlink2  <- wb_dims(x = df, cols = "hLink2", select = "x")
  dims_hlink_2 <- wb_dims(x = df, cols = c("hLink", "hLink2"), select = "x")

  # create workbook
  wb <- wb_workbook()

  wb$add_worksheet()$add_data(x = df)$add_hyperlink(dims = dims_hlink2, target = target, col_names = TRUE)

  exp <- "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://wikipedia.org/wiki/A\" TargetMode=\"External\"/>"
  got <- wb$worksheets_rels[[1]][1]
  expect_equal(exp, got)

  class(df$hLink) <- c(class(df$hLink), "hyperlink")
  # no warning is thrown regarding the hyperlink in the cell, but the hyperlink is from the class, not the
  expect_warning(wb$add_worksheet()$add_data(x = df)$add_hyperlink(target = target), "target not found")
  expect_warning(wb$add_hyperlink(tooltip = target), "tooltip not found")

  exp <- "=HYPERLINK(\"https://r-project.org\")"
  got <- wb$worksheets[[2]]$sheet_data$cc$f[9]
  expect_equal(exp, got)

  # restore plain text
  class(df$hLink) <- c("character")

  # add hyperlink without target
  wb$add_worksheet()$add_data(x = df)$add_hyperlink(dims = wb_dims(x = df, cols = "hLink"))

  got <- "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://r-project.org\" TargetMode=\"External\"/>"
  exp <- wb$worksheets_rels[[3]][2]
  expect_equal(exp, got)

  # FIXME with a typo this happend:
  # Error: `cols` must be an integer or an existing column name of `x`, not hLinkhlink2
  wb$add_worksheet()$add_data_table(x = df)$add_hyperlink(dims = dims_hlink_2, target = target, col_names = TRUE)
  expect_equal(41L, length(wb$worksheets_rels[[4]]))

  wb$add_worksheet()$add_data(x = "noreply@openxlsx2")$add_hyperlink(target = "mailto:noreply@openxlsx2.com", tooltip = "An Invalid E-Mail address")

  exp <- "<hyperlink ref=\"A1\" r:id=\"rId1\" tooltip=\"An Invalid E-Mail address\"/>"
  got <- wb$worksheets[[5]]$hyperlinks
  expect_equal(exp, got)

  wb$add_worksheet()$
    add_data(x = c("https://r-project.org", "https://cran.r-project.org"), dims = "B2:C2")$
    add_hyperlink(dims = "B2:C2", tooltip = c("The R-Project Homepage", "CRAN"))$
    add_data(x = c("https://r-project.org", "https://cran.r-project.org"), dims = "B4:B5")$
    add_hyperlink(dims = "B4:B5", tooltip = c("The R-Project Homepage", "CRAN"))$
    add_data(x = c("The R-Project Homepage", "CRAN"), dims = "B7:B8")$
    add_hyperlink(dims = "B7:B8", target = c("https://r-project.org", "https://cran.r-project.org"))

  exp <- c("<hyperlink ref=\"B2\" r:id=\"rId1\" tooltip=\"The R-Project Homepage\"/>",
           "<hyperlink ref=\"C2\" r:id=\"rId3\" tooltip=\"CRAN\"/>",
           "<hyperlink ref=\"B4\" r:id=\"rId4\" tooltip=\"The R-Project Homepage\"/>",
           "<hyperlink ref=\"B5\" r:id=\"rId5\" tooltip=\"CRAN\"/>",
           "<hyperlink ref=\"B7\" r:id=\"rId6\"/>",
           "<hyperlink ref=\"B8\" r:id=\"rId7\"/>")
  got <- wb$worksheets[[6]]$hyperlinks
  expect_equal(exp, got)

  wb$add_worksheet()$add_data(x = "hyperlinks.xlsx", dims = "B2")$
    add_hyperlink(dims = "B2", target = "/Users/janmarvingarbuszus/Source/openxlsx-data/hyperlinks.xlsx", tooltip = "my testfile")

  exp <- "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"/Users/janmarvingarbuszus/Source/openxlsx-data/hyperlinks.xlsx\" TargetMode=\"External\"/>"
  got <- wb$worksheets_rels[[7]]
  expect_equal(exp, got)

  wb$add_worksheet()$add_data(x = "'Sheet 1'!C5", dims = "B2")$
    add_hyperlink(target = "'Sheet 1'!C5", dims = "B2",
                  tooltip = "An internal reference", is_external = FALSE)$
    add_data(x = "'Sheet 1'!C5", dims = "B4")$
    add_hyperlink(dims = "B4",
                  tooltip = "An internal reference", is_external = FALSE)

  exp <- c(
    "<hyperlink ref=\"B2\" location=\"'Sheet 1'!C5\" display=\"'Sheet 1'!C5\" tooltip=\"An internal reference\"/>",
    "<hyperlink ref=\"B4\" location=\"'Sheet 1'!C5\" tooltip=\"An internal reference\"/>"
  )
  got <- wb$worksheets[[8]]$hyperlinks
  expect_equal(exp, got)

  wb <- wb_workbook(theme = "LibreOffice")$add_worksheet()$
    add_data(x = "https://cran.r-project.org/package=openxlsx2")
  expect_silent(wb$add_hyperlink())

  wb$remove_hyperlink(dims = "A1")
  expect_equal(character(), wb$worksheets[[1]]$hyperlinks)

})
