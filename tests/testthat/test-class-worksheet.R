
test_that("Worksheet Class works", {
  expect_null(assert_worksheet(wb_worksheet()))
})

test_that("test data validation list and sparklines", {

  set.seed(123) # sparklines has a random uri string
  options("openxlsx2_seed" = NULL)

  s1 <- create_sparklines("Sheet 1", "A3:K3", "L3")
  s2 <- create_sparklines("Sheet 1", "A4:K4", "L4")

  wb <- wb_workbook()$
    add_worksheet()$add_data(x = iris[1:30, ])$
    add_worksheet()$add_data(sheet = 2, x = sample(iris$Sepal.Length, 10))$
    add_data_validation(sheet = 1, dims = "A2:A11", type = "list", value = '"O1,O2"')$
    add_sparklines(sheet = 1, sparklines = s1)$
    add_data_validation(sheet = 1, dims = "A12:A21", type = "list", value = '"O2,O3"')$
    add_sparklines(sheet = 1, sparklines = s2)

  exp <- c(
    "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:sparklineGroup displayEmptyCellsAs=\"gap\" xr2:uid=\"{6F57B887-24F1-C14A-942C-4C6EF08E87F7}\"><x14:colorSeries rgb=\"FF376092\"/><x14:colorNegative rgb=\"FFD00000\"/><x14:colorAxis rgb=\"FFD00000\"/><x14:colorMarkers rgb=\"FFD00000\"/><x14:colorFirst rgb=\"FFD00000\"/><x14:colorLast rgb=\"FFD00000\"/><x14:colorHigh rgb=\"FFD00000\"/><x14:colorLow rgb=\"FFD00000\"/><x14:sparklines><x14:sparkline><xm:f>'Sheet 1'!A3:K3</xm:f><xm:sqref>L3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup><x14:sparklineGroup displayEmptyCellsAs=\"gap\" xr2:uid=\"{6F57B887-24F1-C14A-942C-459E3EFAA032}\"><x14:colorSeries rgb=\"FF376092\"/><x14:colorNegative rgb=\"FFD00000\"/><x14:colorAxis rgb=\"FFD00000\"/><x14:colorMarkers rgb=\"FFD00000\"/><x14:colorFirst rgb=\"FFD00000\"/><x14:colorLast rgb=\"FFD00000\"/><x14:colorHigh rgb=\"FFD00000\"/><x14:colorLow rgb=\"FFD00000\"/><x14:sparklines><x14:sparkline><xm:f>'Sheet 1'!A4:K4</xm:f><xm:sqref>L4</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"
  )
  got <- wb$worksheets[[1]]$extLst
  expect_equal(exp, got)

})

test_that("sparkline waivers work", {
  sl <- create_sparklines(dims = "A2:L2", sqref = "M2", markers = "1")

  wb <- wb_workbook()$add_worksheet("Sparklines 1")

  sl_xml <- replace_waiver(sl, wb)

  exp <- "<x14:sparkline><xm:f>'Sparklines 1'!A2:L2</xm:f><xm:sqref>M2</xm:sqref></x14:sparkline>"
  got <- xml_node(sl_xml, "x14:sparklineGroup", "x14:sparklines", "x14:sparkline")
  expect_equal(exp, got)
})

test_that("old and new data validations", {

  temp <- temp_xlsx()

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = sample(c("O1", "O2"), 10, TRUE))$
    add_data(dims = "B1", x = sample(c("O1", "O2"), 10, TRUE))$
    add_data_validation(sheet = 1, dims = "B1:B10", type = "list", value = '"O1,O2"')

  # add data validations list as x14. this was the default in openxlsx and openxlsx2 <= 0.3
  wb$worksheets[[1]]$extLst <- "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\"><x14:dataValidations xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" count=\"2\"><x14:dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\"><x14:formula1><xm:f>\"O1,O2\"</xm:f></x14:formula1><xm:sqref>A2:A11</xm:sqref></x14:dataValidation></x14:dataValidations></ext>"

  wb$save(temp)

  # make sure that it loads
  wb2 <- wb_load(temp)

  # test for equality
  expect_equal(
    wb$worksheets[[1]]$dataValidations,
    wb2$worksheets[[1]]$dataValidations
  )

  expect_equal(
    wb$worksheets[[1]]$extLst,
    wb2$worksheets[[1]]$extLst
  )

})

test_that("set_sheetview", {

  wb <- wb_workbook()$add_worksheet()

  exp <- "<sheetViews><sheetView showGridLines=\"1\" showRowColHeaders=\"1\" tabSelected=\"1\" workbookViewId=\"0\" zoomScale=\"100\"/></sheetViews>"
  got <- wb$worksheets[[1]]$sheetViews
  expect_equal(exp, got)

  exp <- "<sheetViews><sheetView rightToLeft=\"1\" showGridLines=\"1\" showRowColHeaders=\"1\" tabSelected=\"1\" workbookViewId=\"0\" zoomScale=\"100\"/></sheetViews>"

  op <- options("openxlsx2.rightToLeft" = TRUE)
  on.exit(options(op), add = TRUE)
  wb <- wb_workbook()$add_worksheet()

  got <- wb$worksheets[[1]]$sheetViews
  expect_equal(exp, got)

  options("openxlsx2.rightToLeft" = "1")
  wb <- wb_workbook()$add_worksheet()

  got <- wb$worksheets[[1]]$sheetViews

  expect_equal(exp, got)

})

test_that("print options work", {

  temp <- temp_xlsx()

  wb <- wb_workbook() %>%
    wb_add_worksheet(grid_lines = FALSE) %>%
    wb_add_data(x = iris) %>%
    wb_add_worksheet(grid_lines = TRUE) %>%
    wb_add_data(x = mtcars)

  exp <- character()
  got <- wb$worksheets[[1]]$printOptions
  expect_equal(exp, got)

  exp <- "<printOptions gridLines=\"1\" gridLinesSet=\"1\"/>"
  got <- wb$worksheets[[2]]$printOptions
  expect_equal(exp, got)

  wb$save(temp)
  wb <- wb_load(temp)

  got <- wb$worksheets[[2]]$printOptions
  expect_equal(exp, got)

})

test_that("ignore_error works", {

  wb <- wb_workbook()$add_worksheet()
  wb$add_data(dims = "B1", x = t(c(1, 2, 3)), col_names = FALSE)
  wb$add_formula(dims = "A1", x = "SUM(B1:C1)")

  # 1
  wb$worksheets[[1]]$ignore_error(dims = "A1", formulaRange = TRUE)

  exp <- "<ignoredErrors><ignoredError sqref=\"A1\" formulaRange=\"1\"/></ignoredErrors>"
  got <- wb$worksheets[[1]]$ignoredErrors
  expect_equal(exp, got)

  # 2
  wb$worksheets[[1]]$ignore_error(dims = "A2", calculatedColumn = TRUE, emptyCellReference = TRUE, evalError = TRUE, formula = TRUE, formulaRange = TRUE, listDataValidation = TRUE, numberStoredAsText = TRUE, twoDigitTextYear = TRUE, unlockedFormula = TRUE)

  exp <- "<ignoredErrors><ignoredError formulaRange=\"1\" sqref=\"A1\"/><ignoredError formulaRange=\"1\" sqref=\"A2\" calculatedColumn=\"1\" emptyCellReference=\"1\" evalError=\"1\" formula=\"1\" listDataValidation=\"1\" numberStoredAsText=\"1\" twoDigitTextYear=\"1\" unlockedFormula=\"1\"/></ignoredErrors>"
  got <- wb$worksheets[[1]]$ignoredErrors
  expect_equal(exp, got)

})

test_that("tab_color works", {

  # worksheet
  wb <- wb_workbook()$
    add_worksheet(tab_color = "red")$
    add_worksheet(tab_color = wb_color("red"))

  expect_equal(
    wb$worksheets[[1]]$sheetPr,
    wb$worksheets[[2]]$sheetPr
  )

  # chartsheet
  wb <- wb_workbook()$
    add_chartsheet(tab_color = "red")$
    add_chartsheet(tab_color = wb_color("red"))

  expect_equal(
    wb$worksheets[[1]]$sheetPr,
    wb$worksheets[[2]]$sheetPr
  )

  # use color theme
  wb <- wb_workbook()$
    add_worksheet(tab_color = wb_color(theme = 4))$
    add_chartsheet(tab_color = wb_color(theme = 4))

  expect_equal(
    wb$worksheets[[1]]$sheetPr,
    wb$worksheets[[2]]$sheetPr
  )

  # error with invalid tab_color. blau is German for blue.
  expect_error(
    wb <- wb_workbook()$
      add_worksheet(tab_color = "blau"),
    "Invalid tab_color in add_worksheet"
  )
  expect_error(
    wb <- wb_workbook()$
      add_chartsheet(tab_color = "blau"),
    "Invalid tab_color in add_chartsheet"
  )

})

test_that("setting and loading header/footer attributes works", {
  wb <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_set_header_footer(
      header             = c(NA, "Header", NA),
      scale_with_doc     = TRUE,
      align_with_margins = TRUE
    ) %>%
    wb_page_setup(orientation = "landscape", fit_to_width = 1) %>%
    wb_set_sheetview(view = "pageLayout", zoom_scale = 40) %>%
    wb_add_data(x = as.data.frame(matrix(1:500, ncol = 25)))

  temp <- temp_xlsx()
  wb$save(temp)
  rm(wb)

  wb <- wb_load(temp)
  expect_true(wb$worksheets[[1]]$scale_with_doc)
  expect_true(wb$worksheets[[1]]$align_with_margins)
})

test_that("updating page header / footer works", {
  wb <- wb_workbook()$add_worksheet()$set_sheetview(view = "pageLayout")
  wb$add_data(x = matrix(1, nrow = 150, ncol = 1))

  first_hf <- wb$worksheets[[1]]$headerFooter

  wb$set_header_footer(
    header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
    footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
    even_header = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
    even_footer = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
    first_header = c("TOP", "OF FIRST", "PAGE"),
    first_footer = c("BOTTOM", "OF FIRST", "PAGE")
  )

  second_hf <- wb$worksheets[[1]]$headerFooter

  wb$set_header_footer(
    header = NA,
    footer = NA,
    even_header = NA,
    even_footer = NA,
    first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
    first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
  )

  third_hf <- wb$worksheets[[1]]$headerFooter

  wb$set_header_footer(
    first_header = c("FIRST ONLY L", NA, "FIRST ONLY R"),
    first_footer = c("FIRST ONLY L", NA, "FIRST ONLY R")
  )

  fourth_hf <- wb$worksheets[[1]]$headerFooter


  expect_equal(NULL, first_hf)


  exp <- list(oddHeader = list("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
              oddFooter = list("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
              evenHeader = list("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
              evenFooter = list("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
              firstHeader = list("TOP", "OF FIRST", "PAGE"),
              firstFooter = list("BOTTOM", "OF FIRST", "PAGE"))
  expect_equal(exp, second_hf)


  exp <- list(oddHeader = list("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
              oddFooter = list("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
              evenHeader = list("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
              evenFooter = list("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
              firstHeader = list("FIRST ONLY L", "OF FIRST", "FIRST ONLY R"),
              firstFooter = list("FIRST ONLY L", "OF FIRST", "FIRST ONLY R"))
  expect_equal(exp, third_hf)


  exp <- list(oddHeader = list(),
              oddFooter = list(),
              evenHeader = list(),
              evenFooter = list(),
              firstHeader = list("FIRST ONLY L", "OF FIRST", "FIRST ONLY R"),
              firstFooter = list("FIRST ONLY L", "OF FIRST", "FIRST ONLY R"))
  expect_equal(exp, fourth_hf)

  expect_error(wb$set_header_footer(header = c("foo", "bar")), "must have length 3 where elements correspond to positions: left, center, right.")
})

test_that("set_grid_lines() works", {

  wb <- wb_workbook()

  exp <- "<sheetViews><sheetView showGridLines=\"0\" showRowColHeaders=\"1\" tabSelected=\"1\" workbookViewId=\"0\" zoomScale=\"100\"/></sheetViews>"
  wb$add_worksheet(grid_lines = FALSE)
  got <- wb$worksheets[[1]]$sheetViews
  expect_equal(exp, got)

  # Show grid lines
  wb$set_grid_lines(print = FALSE, show = TRUE)
  exp <- "<sheetViews><sheetView showGridLines=\"1\" showRowColHeaders=\"1\" tabSelected=\"1\" workbookViewId=\"0\" zoomScale=\"100\"/></sheetViews>"
  got <- wb$worksheets[[1]]$sheetViews
  expect_equal(exp, got)

  exp <- "<sheetViews><sheetView showGridLines=\"0\" showRowColHeaders=\"1\" tabSelected=\"1\" workbookViewId=\"0\" zoomScale=\"100\"/></sheetViews>"
  wb$set_grid_lines(print = FALSE, show = FALSE)
  got <- wb$worksheets[[1]]$sheetViews
  expect_equal(exp, got)
  wb$worksheets[[1]]$sheetViews

  # Show grid lines
  wb$set_grid_lines(print = TRUE, show = FALSE)
  exp <- "<printOptions gridLines=\"1\" gridLinesSet=\"1\"/>"
  got <- wb$worksheets[[1]]$printOptions
  expect_equal(exp, got)

  exp <- "<printOptions gridLines=\"0\" gridLinesSet=\"0\"/>"
  wb$set_grid_lines(print = FALSE, show = FALSE)
  got <- wb$worksheets[[1]]$printOptions
  expect_equal(exp, got)

})
