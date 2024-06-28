
testsetup()

test_that("Workbook class", {
  expect_null(assert_workbook(wb_workbook()))
})


test_that("wb_set_col_widths works", {
# TODO use wb$wb_set_col_widths()

  wb <- wbWorkbook$new()
  wb$add_worksheet("test")
  wb$add_data("test", mtcars)

  # set column width to 12
  expect_silent(wb$set_col_widths("test", widths = 12L, cols = seq_along(mtcars)))
  expect_equal(
    wb$worksheets[[1]]$cols_attr,
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"12.711\"/>"
  )

  # wrong sheet
  expect_error(wb$set_col_widths("test2", widths = 12L, cols = seq_along(mtcars)))

  # reset the column with, we do not provide an option ot remove the column entry
  expect_silent(wb$set_col_widths("test", cols = seq_along(mtcars)))
  expect_equal(
    wb$worksheets[[1]]$cols_attr,
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"9.141\"/>"
  )

  # create column width for column 25
  expect_silent(wb$set_col_widths("test", cols = "Y", widths = 22))
  expect_equal(
    c("<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"9.141\"/>",
      "<col min=\"12\" max=\"24\" width=\"8.43\"/>",
      "<col min=\"25\" max=\"25\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"22.711\"/>"),
    wb$worksheets[[1]]$cols_attr
  )

  wb <- wb_workbook()$
    add_worksheet()$
    set_col_widths(cols = 1:10, widths = (8:17) + .5)$
    add_data(x = rbind(8:17), col_names = FALSE)

  exp <- c(
    "<col min=\"1\" max=\"1\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"9.211\"/>",
    "<col min=\"2\" max=\"2\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"10.211\"/>",
    "<col min=\"3\" max=\"3\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"11.211\"/>",
    "<col min=\"4\" max=\"4\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"12.211\"/>",
    "<col min=\"5\" max=\"5\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"13.211\"/>",
    "<col min=\"6\" max=\"6\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"14.211\"/>",
    "<col min=\"7\" max=\"7\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"15.211\"/>",
    "<col min=\"8\" max=\"8\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"16.211\"/>",
    "<col min=\"9\" max=\"9\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"17.211\"/>",
    "<col min=\"10\" max=\"10\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"18.211\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)

  wb <- wb_workbook()$add_worksheet()
  wb$worksheets[[1]]$cols_attr <- c(
    "<col min=\"1\" max=\"17\" width=\"30.77734375\" style=\"16\" customWidth=\"1\"/>",
    "<col min=\"18\" max=\"16384\" width=\"8.88671875\" style=\"16\"/>"
  )

  expect_silent(wb$set_col_widths(cols = 19, widths = 9))

})

test_that("set_col_widths informs when inconsistent lengths are supplied", {
  wb <- wbWorkbook$new()
  wb$add_worksheet("test")

  expect_warning(wb$set_col_widths(cols = c(1, 2, 3), widths = c(2, 3)), "compatible length")
  expect_error(wb$set_col_widths(cols = "Y", widths = 1:2), "More widths than column")
  expect_error(wb$set_col_widths("test", cols = "Y", hidden = 1:2), "hidden argument is longer")
  expect_warning(wb$set_col_widths(cols = c("X", "Y", "Z"), hidden = c(1, 0)), "compatible length")
})

test_that("option maxWidth works", {

  op <- options("openxlsx2.maxWidth" = 6)
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(
      x = paste0(letters, collapse = ""),
      y = paste0(letters, collapse = "")
  ))$set_col_widths(cols = 1:2, widths = "auto")

  exp <- "<col min=\"1\" max=\"2\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"6\"/>"
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)

})


# order -------------------------------------------------------------------

test_that("$set_order() works", {
  wb <- wb_workbook()
  wb$add_worksheet("a")
  wb$add_worksheet("b")
  wb$add_worksheet("c")

  expect_identical(wb$sheetOrder, 1:3)
  exp <- letters[1:3]
  names(exp) <- exp
  expect_identical(wb$get_sheet_names(), exp)

  wb$set_order(3:1)
  expect_identical(wb$sheetOrder, 3:1)
  exp <- letters[3:1]
  names(exp) <- exp
  expect_identical(wb$get_sheet_names(), exp)
})


# sheet names -------------------------------------------------------------

test_that("$set_sheet_names() and $get_sheet_names() work", {
  wb <- wb_workbook()$add_worksheet()$add_worksheet()
  wb$set_sheet_names(new = c("a", "b & c"))

  # return a names character vector
  res <- wb$get_sheet_names()
  exp <- c(a = "a", "b & c" = "b & c")
  expect_identical(res, exp)

  # return a names character vector
  res <- wb$get_sheet_names(escape = TRUE)
  exp <- c(a = "a", "b & c" = replace_legal_chars("b & c"))
  expect_identical(res, exp)

  # should be able to check the original values, too
  res <- wb$.__enclos_env__$private$get_sheet_index("b & c")
  expect_identical(res, 2L)

  # make sure that it works silently
  wb <- wb_load(file = system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2"))
  expect_silent(wb$set_sheet_names(old = "SUM", new = "Sheet 1"))

  exp <- c(`Sheet 1` = "Sheet 1")
  got <- wb$get_sheet_names()
  expect_equal(exp, got)
})

# data validation ---------------------------------------------------------


test_that("data validation", {

  temp <- temp_xlsx()

  df <- data.frame(
    "d" = as.Date("2016-01-01") + -5:5,
    "t" = as.POSIXct("2016-01-01") + -5:5 * 10000
  )

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data_table(x = iris)$
    # whole numbers are fine
    add_data_validation(dims = "A2:C151", type = "whole",
                        operator = "between", value = c(1, 9)
    )$
    # text width 7-9 is fine
    add_data_validation(dims = "E2:E151", type = "textLength",
                        operator = "between", value = c(7, 9)
    )$
    ## Date and Time cell validation
    add_worksheet("Sheet 2")$
    add_data_table(x = df)$
    # date >= 2016-01-01 is fine
    add_data_validation(dims = "A2:A12", type = "date",
                        operator = "greaterThanOrEqual", value = as.Date("2016-01-01")
    )$
    # a few timestamps are fine
    add_data_validation(dims = "B2:B12", type = "time",
                        operator = "between", value = df$t[c(4, 8)]
    )$
    ## validate list: validate inputs on one sheet with another
    add_worksheet("Sheet 3")$
    add_data_table(x = iris[1:30, ])$
    add_worksheet("Sheet 4")$
    add_data(x = sample(iris$Sepal.Length, 10))$
    add_data_validation("Sheet 3", dims = "A2:A31", type = "list",
                        value = "'Sheet 4'!$A$1:$A$10")

  exp <- c(
    "<dataValidation type=\"whole\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:C151\"><formula1>1</formula1><formula2>9</formula2></dataValidation>",
    "<dataValidation type=\"textLength\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"E2:E151\"><formula1>7</formula1><formula2>9</formula2></dataValidation>"
  )
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)


  exp <- c(
    "<dataValidation type=\"date\" operator=\"greaterThanOrEqual\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A12\"><formula1>42370</formula1></dataValidation>",
    "<dataValidation type=\"time\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B2:B12\"><formula1>42369.7685185185</formula1><formula2>42370.2314814815</formula2></dataValidation>"
  )
  got <- wb$worksheets[[2]]$dataValidations
  expect_equal(exp, got)


  exp <- c(
    "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A31\"><formula1>'Sheet 4'!$A$1:$A$10</formula1></dataValidation>"
  )
  got <- wb$worksheets[[3]]$dataValidations
  expect_equal(exp, got)

  wb$save(temp)

  wb2 <- wb_load(temp)

  # wb2$add_data_validation("Sheet 3", col = 2, rows = 2:31, type = "list",
  #                         value = "'Sheet 4'!$A$1:$A$10")
  # wb2$save(temp)

  expect_equal(
    wb$worksheets[[1]]$dataValidations,
    wb2$worksheets[[1]]$dataValidations
  )

  expect_equal(
    wb$worksheets[[2]]$dataValidations,
    wb2$worksheets[[2]]$dataValidations
  )

  expect_equal(
    wb$worksheets[[3]]$dataValidations,
    wb2$worksheets[[3]]$dataValidations
  )

  expect_warning(
    wb2$add_data_validation("Sheet 3", cols = 2, rows = 2:31, type = "list",
                            value = "'Sheet 4'!$A$1:$A$10"),
    "'cols/rows' is deprecated."
  )

  exp <- c(
    "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A31\"><formula1>'Sheet 4'!$A$1:$A$10</formula1></dataValidation>",
    "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B2:B31\"><formula1>'Sheet 4'!$A$1:$A$10</formula1></dataValidation>"
  )
  got <- wb2$worksheets[[3]]$dataValidations
  expect_equal(exp, got)

  ### tests if conditions

  # test col2int
  expect_warning(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      # whole numbers are fine
      add_data_validation(cols = "A", rows = 2:151, type = "whole",
                          operator = "between", value = c(1, 9)
      ),
    "'cols/rows' is deprecated."
  )

  exp <- "<dataValidation type=\"whole\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A151\"><formula1>1</formula1><formula2>9</formula2></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)


  # to many values
  expect_error(
    wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data_table(x = head(iris))$
    add_data_validation(dims = "A2:A151", type = "whole",
                        operator = "between", value = c(1, 9, 19)
    ),
    "length <= 2"
  )

  # wrong type
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      add_data_validation(dims = "A2:A151", type = "even",
                          operator = "between", value = c(1, 9)
      ),
    "Invalid 'type' argument!"
  )

  # wrong operator
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      add_data_validation(dims = "A2:A151", type = "whole",
                          operator = "lower", value = c(1, 9)
      ),
    "Invalid 'operator' argument!"
  )

  # wrong value for date
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      # whole numbers are fine
      add_data_validation(dims = "A2:A12", type = "date",
                          operator = "greaterThanOrEqual", value = 7
      ),
    "If type == 'date' value argument must be a Date vector"
  )

  # wrong value for time
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      # whole numbers are fine
      add_data_validation(dims = "A2:A12", type = "time",
                          operator = "greaterThanOrEqual", value = 7
      ),
    "If type == 'time' value argument must be a POSIXct or POSIXlt vector."
  )


  # some more options
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = c(-1:1), colNames = FALSE)$
    # whole numbers are fine
    add_data_validation(dims = "A1:A3", type = "whole",
                        operator = "greaterThan", value = c(0),
                        errorStyle = "information", errorTitle = "ERROR!",
                        error = "Some error ocurred!",
                        promptTitle = "PROMPT!",
                        prompt = "Choose something!"
    )

  exp <- "<dataValidation type=\"whole\" operator=\"greaterThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1:A3\" errorStyle=\"information\" errorTitle=\"ERROR!\" error=\"Some error ocurred!\" promptTitle=\"PROMPT!\" prompt=\"Choose something!\"><formula1>0</formula1></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)

  # add custom data
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = data.frame(x = 1, y = 2), colNames = FALSE)$
    # whole numbers are fine
    add_data_validation(dims = "A1:A3", type = "custom", value = "A1=B1")

  exp <- "<dataValidation type=\"custom\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1:A3\"><formula1>A1=B1</formula1></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)

})


test_that("clone worksheet", {

  ## Dummy tests - not sure how to test these from R ##

  # # clone chartsheet ----------------------------------------------------
  fl <- testfile_path("mtcars_chart.xlsx")
  wb <- wb_load(fl)
  # wb$get_sheet_names() # chartsheet has no named name?
  expect_silent(wb$clone_worksheet(1, "Clone 1"))
  expect_s3_class(wb$worksheets[[5]], "wbChartSheet")
  # wb$open()

  # clone pivot table and drawing -----------------------------------------
  fl <- testfile_path("loadExample.xlsx")
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet(4, "Clone 1"))

  # sheets 4 & 5 both reference the same pivot table in different drawing
  # once the file is opened, both pivot tables behave independently
  exp <- c(
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable2.xml\"/>",
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing5.xml\"/>"
  )
  got <- wb$worksheets_rels[[5]]
  expect_equal(exp, got)
  # wb$open()

  # clone drawing ---------------------------------------------------------
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet("testing", "Clone1"))

  expect_false(identical(wb$worksheets_rels[2], wb$worksheets_rels[5]))
  # wb$open()

  # clone sheet with table ------------------------------------------------
  fl <- testfile_path("tableStyles.xlsx")
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet(1, "clone"))

  expect_false(identical(wb$tables$tab_xml[1], wb$tables$tab_xml[2]))
  # wb$open()

  # clone sheet with chart ------------------------------------------------
  fl <- testfile_path("mtcars_chart.xlsx")
  wb <- wb_load(fl)
  wb$clone_worksheet(2, "Clone 1")

  expect_match(wb$charts$chart[2], "test")
  expect_match(wb$charts$chart[3], "'Clone 1'")
  # wb$open()

  # clone slicer ----------------------------------------------------------
  fl <- testfile_path("loadExample.xlsx")
  wb <- wb_load(fl)
  expect_warning(wb$clone_worksheet("IrisSample", "Clone1"),
                 "Cloning slicers is not yet supported. It will not appear on the sheet.")
  # wb$open()

})

test_that("set and remove row heights work", {

  ## add row heights
  wb <- wb_workbook()$
    add_worksheet()$
    set_row_heights(
      rows = c(1, 4, 22, 2, 19),
      heights = c(24, 28, 32, 42, 33)
    )

  exp <- structure(
    list(
      customHeight = c("1", "1", "1", "1", "1"),
      ht = c("24", "42", "28", "33", "32"),
      r = c("1", "2", "4", "19", "22")
    ),
    row.names = c(1L, 2L, 4L, 19L, 22L),
    class = "data.frame"
  )
  got <- wb$worksheets[[1]]$sheet_data$row_attr[c(1, 2, 4, 19, 22), c("customHeight", "ht", "r")]
  expect_equal(exp, got)

  ## remove row heights
  wb$remove_row_heights(rows = 1:21)
  exp <- structure(
    list(
      customHeight = c("", "", "", "", "1"),
      ht = c("", "", "", "", "32"),
      r = c("1", "2", "4", "19", "22")
    ),
    row.names = c(1L, 2L, 4L, 19L, 22L),
    class = "data.frame"
  )
  got <- wb$worksheets[[1]]$sheet_data$row_attr[c(1, 2, 4, 19, 22), c("customHeight", "ht", "r")]
  expect_equal(exp, got)

  expect_warning(
    wb$add_worksheet()$remove_row_heights(rows = 1:3),
    "There are no initialized rows on this sheet"
  )

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)$
    set_row_heights(rows = 5:15, hidden = TRUE)

  exp <- structure(
    c(22L, `1` = 11L),
    dim = 2L,
    dimnames = structure(
      list(
        c("", "1")
      ),
      names = ""
    ),
    class = "table"
  )
  got <- table(wb$worksheets[[1]]$sheet_data$row_attr$hidden)

  expect_equal(exp, got)

})

test_that("add_drawing works", {

  skip_if_not_installed("rvg")
  skip_if_not_installed("ggplot2")

  require(rvg)
  require(ggplot2)

  tmp <- tempfile(fileext = "drawing.xml")

  ## rvg example
  dml_xlsx(file =  tmp, fonts = list(sans = "Bradley Hand"))
  print(
    ggplot(data = iris,
           mapping = aes(x = Sepal.Length, y = Petal.Width)) +
      geom_point() + labs(title = "With font Bradley Hand") +
      theme_minimal(base_family = "sans", base_size = 18)
  )
  dev.off()

  wb <- wb_workbook()$
    add_worksheet()$
    add_drawing(xml = tmp)$
    add_drawing(xml = tmp, dims = "A1:H10")$
    add_drawing(xml = tmp, dims = "L1")$
    add_drawing(xml = tmp, dims = NULL)$
    add_drawing(xml = tmp, dims = "L19")

  expect_length(wb$drawings, 1L)

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  expect_silent(wb$add_image(file = img))

})

test_that("add_drawing works", {

  skip_if_not_installed("mschart")

  require(mschart)

  # write data starting at B2
  wb <- wb_workbook()$add_worksheet()$
    add_data(x = mtcars, dims = "B2")$
    add_data(x = data.frame(name = rownames(mtcars)), dims = "A2")

  # create wb_data object this will tell this mschart from this PR to create a file corresponding to openxlsx2
  dat <- wb_data(wb, 1)
  expect_equal(c(32L, 12L), dim(dat))

  dat <- wb_data(wb, 1, dims = "A2:G6")

  exp <- structure(
    list(
      name = c("Mazda RX4", "Mazda RX4 Wag", "Datsun 710", "Hornet 4 Drive"),
      mpg = c(21, 21, 22.8, 21.4),
      cyl = c(6, 6, 4, 6),
      disp = c(160, 160, 108, 258),
      hp = c(110, 110, 93, 110),
      drat = c(3.9, 3.9, 3.85, 3.08),
      wt = c(2.62, 2.875, 2.32, 3.215)
    ),
    row.names = 3:6,
    class = c("wb_data", "data.frame"),
    dims = structure(
      list(
        A = c("A2", "A3", "A4", "A5", "A6"),
        B = c("B2", "B3", "B4", "B5", "B6"),
        C = c("C2", "C3", "C4", "C5", "C6"),
        D = c("D2", "D3", "D4", "D5", "D6"),
        E = c("E2", "E3", "E4", "E5", "E6"),
        F = c("F2", "F3", "F4", "F5", "F6"),
        G = c("G2", "G3", "G4", "G5", "G6")
      ),
      row.names = 2:6,
      class = "data.frame"),
    sheet = "Sheet 1")

  expect_equal(exp, dat)

  # call ms_scatterplot
  scatter_plot <- ms_scatterchart(
    data = dat,
    x = "mpg",
    y = c("disp", "hp"),
    labels = c("disp", "hp")
  )

  # add the scatterplots to the data
  wb <- wb %>%
    wb_add_mschart(dims = "F4:L20", graph = scatter_plot)

  expect_equal(NROW(wb$charts), 1L)

  chart_01 <- ms_linechart(
    data = us_indus_prod,
    x = "date", y = "value",
    group = "type"
  )

  wb$add_worksheet()
  wb$add_mschart(dims = "F4:L20", graph = chart_01)

  exp <- list(
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart1.xml\"/>",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart2.xml\"/>"
  )
  got <- wb$drawings_rels
  expect_equal(exp, got)


  # write data starting at B2
  wb <- wb_workbook()$
    add_worksheet()$add_data(x = mtcars)$
    add_worksheet()$add_data(x = mtcars)$
    add_worksheet()$add_data(x = mtcars)$
    add_worksheet()$add_data(x = mtcars)$
    add_mschart(dims = "F4:L20", 2, graph = chart_01)$
    add_mschart(dims = "F4:L20", 3, graph = chart_01)

  exp <- list(
    character(0),
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing2.xml\"/>",
    character(0)
  )
  got <- wb$worksheets_rels
  expect_equal(exp, got)

  ## write different anchors
  wb <- wb_workbook()$
    add_worksheet()$add_data(x = mtcars)

  scatter_plot <- ms_scatterchart(
    data = wb_data(wb),
    x = "mpg",
    y = c("disp", "hp")
  )

  wb$
    add_mschart(graph = scatter_plot)$
    add_mschart(dims = "A1", graph = scatter_plot)$
    add_mschart(dims = "F4:L20", graph = scatter_plot)

  drawings <- as.character(wb$drawings)
  expect_match(drawings, "absoluteAnchor")
  expect_match(drawings, "oneCellAnchor")
  expect_match(drawings, "twoCellAnchor")

})

test_that("add_chartsheet works", {

  skip_if_not_installed("mschart")

  require(mschart)

  wb <- wb_workbook()$
    add_worksheet("A & B")$
    add_data(x = mtcars)$
    add_chartsheet(tabColour = "red")

  dat <- wb_data(wb, 1, dims = "A1:E6")

  # call ms_scatterplot
  data_plot <- ms_scatterchart(
    data = dat,
    x = "mpg",
    y = c("disp", "hp"),
    labels = c("disp", "hp")
  )

  wb$add_mschart(graph = data_plot)

  expect_equal(nrow(wb$charts), 1)

  expect_match(wb$charts$chart, "A &amp; B")

  expect_true(wb$is_chartsheet[[2]])

  # add new worksheet and replace chart on chartsheet
  wb$add_worksheet()$add_data(x = mtcars)
  dat <- wb_data(wb, dims = "A1:E1;A7:E15")
  data_plot <- ms_scatterchart(
    data = dat,
    x = "mpg",
    y = c("disp", "hp"),
    labels = c("disp", "hp")
  )
  wb$add_mschart(sheet = 2, graph = data_plot)

  expect_equal(nrow(wb$charts), 2L)

  exp <- "xdr:absoluteAnchor"
  got <- xml_node_name(unlist(wb$drawings), "xdr:wsDr")
  expect_equal(exp, got)

})

test_that("multiple charts on a sheet work as expected", {

  skip_if_not_installed("mschart")

  require(mschart)

  ## Add mschart to worksheet (adds data and chart)
  scatter <- ms_scatterchart(
    data = iris,
    x = "Sepal.Length",
    y = "Sepal.Width",
    group = "Species"
  )
  scatter <- chart_settings(scatter, scatterstyle = "marker")

  wb <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_mschart(dims = "F4:L20", graph = scatter) %>%
    wb_add_mschart(dims = "F24:L40", graph = scatter) %>%
    wb_add_worksheet() %>%
    wb_add_mschart(dims = "F4:L20", graph = scatter) %>%
    wb_add_mschart(dims = "F24:L40", graph = scatter)

  exp <- c(TRUE, TRUE)
  got <- grepl(pattern = "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/>", wb$drawings)
  expect_equal(exp, got)

  got <- grepl(pattern = "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId2\"/>", wb$drawings)
  expect_equal(exp, got)

})

test_that("various image functions work as expected", {

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  wb <- wb_workbook()$
    add_worksheet()$
    add_image(file = img, width = 6, height = 5, dims = NULL)$
    add_worksheet()$
    add_image(dims = "B2", file = img, rowOffset = 90000, colOffset = 90000)$
    add_worksheet()$
    add_image(dims = "B2:K8", file = img)

  exp <- c("xdr:absoluteAnchor", "xdr:oneCellAnchor", "xdr:twoCellAnchor")
  got <- wb$drawings %>% xml_node_name(level1 = "xdr:wsDr")
  expect_equal(exp, got)

  exp <- "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>90000</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>90000</xdr:rowOff></xdr:from>"
  got <- wb$drawings[[2]] %>% xml_node("xdr:wsDr", "xdr:oneCellAnchor", "xdr:from")
  expect_equal(exp, got)

  exp <- "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
  got <- wb$drawings[[3]] %>% xml_node("xdr:wsDr", "xdr:twoCellAnchor", "xdr:from")
  expect_equal(exp, got)

  exp <- "<xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
  got <- wb$drawings[[3]] %>% xml_node("xdr:wsDr", "xdr:twoCellAnchor", "xdr:to")
  expect_equal(exp, got)

  expect_warning(
    wb$add_worksheet()$add_image(file = img, width = 6, height = 5, dims = NULL, startRow = 2, startCol = 2),
    "'start_col/start_row' is deprecated."
  )

})

test_that("image relships work with comment", {

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")

  c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
  wb$add_comment(dims = "B12", comment = c1)

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

  wb$add_image("Sheet 1", dims = "C5", file = img, width = 6, height = 5)

  exp <- "<drawing r:id=\"rId3\"/>"
  got <- wb$worksheets[[1]]$drawing
  expect_equal(exp, got)

})

test_that("start_col/start_row works as expected", {
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  expect_warning(wb <- wb_workbook()$add_worksheet()$add_image(file = img, start_row = 5), "'start_col/start_row' is deprecated.")

  exp <- "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
  got <- xml_node(wb$drawings[[1]], "xdr:wsDr", "xdr:oneCellAnchor", "xdr:from")
  expect_equal(exp, got)

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  expect_warning(wb <- wb_workbook()$add_worksheet()$add_image(file = img, startRow = 5), "'start_col/start_row' is deprecated.")

  exp <- "<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
  got <- xml_node(wb$drawings[[1]], "xdr:wsDr", "xdr:oneCellAnchor", "xdr:from")
  expect_equal(exp, got)


  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  expect_warning(wb <- wb_workbook()$add_worksheet()$add_image(file = img, start_col = 5), "'start_col/start_row' is deprecated.")

  exp <- "<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
  got <- xml_node(wb$drawings[[1]], "xdr:wsDr", "xdr:oneCellAnchor", "xdr:from")
  expect_equal(exp, got)

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  expect_warning(wb <- wb_workbook()$add_worksheet()$add_image(file = img, startCol = 5), "'start_col/start_row' is deprecated.")

  exp <- "<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
  got <- xml_node(wb$drawings[[1]], "xdr:wsDr", "xdr:oneCellAnchor", "xdr:from")
  expect_equal(exp, got)


  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  expect_warning(wb <- wb_workbook()$add_worksheet()$add_image(file = img, start_col = 5, start_row = 5), "'start_col/start_row' is deprecated.")

  exp <- "<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
  got <- xml_node(wb$drawings[[1]], "xdr:wsDr", "xdr:oneCellAnchor", "xdr:from")
  expect_equal(exp, got)


  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  expect_warning(wb <- wb_workbook()$add_worksheet()$add_image(file = img, startCol = 5, startRow = 5), "'start_col/start_row' is deprecated.")

  exp <- "<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
  got <- xml_node(wb$drawings[[1]], "xdr:wsDr", "xdr:oneCellAnchor", "xdr:from")
  expect_equal(exp, got)
})

test_that("workbook themes work", {

  wb <- wb_workbook()$add_worksheet()
  exp <- "Aptos Narrow"
  got <- wb$get_base_font()$name$val
  expect_equal(exp, got)

  wb <- wb_workbook(theme = "Office 2013 - 2022 Theme")$add_worksheet()
  exp <- "Calibri"
  got <- wb$get_base_font()$name$val
  expect_equal(exp, got)

  wb <- wb_workbook(theme = "Old Office Theme")$add_worksheet()
  exp <- "Calibri"
  got <- wb$get_base_font()$name$val
  expect_equal(exp, got)

  wb <- wb_workbook(theme = 1)$add_worksheet()
  exp <- "Rockwell"
  got <- wb$get_base_font()$name$val
  expect_equal(exp, got)

  expect_message(
    wb <- wb_workbook(theme = "Foo")$add_worksheet(),
    "theme Foo not found falling back to default theme"
  )
  exp <- "Aptos Narrow"
  got <- wb$get_base_font()$name$val
  expect_equal(exp, got)

})

test_that("changing sheet names works with named regions", {

  filename <- testfile_path("namedRegions2.xlsx")
  wb <- wb_load(filename)

  wb$set_sheet_names("Sheet1", "new name")
  wb$set_sheet_names("Sheet with space", "Sheet_without_space")

  exp <- c(
    "<definedName name=\"barref\" localSheetId=\"0\">'Sheet_without_space'!$B$4</definedName>",
    "<definedName name=\"barref\" localSheetId=\"1\">'new name'!$B$4</definedName>"
  )
  got <- wb$workbook$definedNames[seq_len(2)]
  expect_equal(exp, got)

})

test_that("numfmt in pivot tables works", {

  ## example code
  df <- data.frame(
    Plant = c("A", "C", "C", "B", "B", "C", "C", "C", "A", "C"),
    Location = c("E", "F", "E", "E", "F", "E", "E", "G", "E", "F"),
    Status = c("good", "good", "good", "good", "good", "good", "good", "good", "good", "bad"),
    Units = c(0.95, 0.95, 0.95, 0.95, 0.89, 0.89, 0.94, 0.94, 0.9, 0.9),
    stringsAsFactors = FALSE
  )

  ## Create the workbook and the pivot table
  wb <- wb_workbook()$
    add_worksheet("Data")$
    add_data(x = df, startCol = 1, startRow = 2)

  df <- wb_data(wb, 1, dims = "A2:D10")
  wb$
    add_pivot_table(df, dims = "A3", rows = "Plant",
                    filter = c("Location", "Status"), data = "Units")$
    add_pivot_table(df, dims = "A3", rows = "Plant",
                    filter = c("Location", "Status"), data = "Units",
                    params = list(numfmt = c(formatCode = "#,###0"), sort_row = "ascending"))$
    add_pivot_table(df, dims = "A3", rows = "Plant",
                    filter = c("Location", "Status"), data = "Units",
                    params = list(numfmt = c(numfmt = 10), sort_row = "descending"))

  exp <- c(
    "<dataField name=\"Sum of Units\" fld=\"3\" baseField=\"0\" baseItem=\"0\"/>",
    "<dataField name=\"Sum of Units\" fld=\"3\" baseField=\"0\" baseItem=\"0\" numFmtId=\"165\"/>",
    "<dataField name=\"Sum of Units\" fld=\"3\" baseField=\"0\" baseItem=\"0\" numFmtId=\"10\"/>"
  )
  got <- xml_node(wb$pivotTables, "pivotTableDefinition", "dataFields", "dataField")
  expect_equal(exp, got)

  ## sort by column and row
  df <- mtcars

  ## Create the workbook and the pivot table
  wb <- wb_workbook()$
    add_worksheet("Data")$
    add_data(x = df, start_col = 1, start_row = 2)

  df <- wb_data(wb)
  wb$add_pivot_table(df, dims = "A3", rows = "cyl", cols = "gear",
                     data = c("vs", "am"), params = list(sort_row = 1, sort_col = -2))

  wb$add_pivot_table(df, dims = "A3", rows = "gear",
                     filter = c("cyl"), data = c("vs", "am"),
                     params = list(sort_row = "descending"))


  exp <- c(
    "<pivotField axis=\"axisRow\" showAll=\"0\" sortType=\"ascending\"><items count=\"4\"><item x=\"1\"/><item x=\"0\"/><item x=\"2\"/><item t=\"default\"/></items><autoSortScope><pivotArea dataOnly=\"0\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\"><x v=\"0\"/></reference></references></pivotArea></autoSortScope></pivotField>",
    "<pivotField axis=\"axisCol\" showAll=\"0\" sortType=\"descending\"><items count=\"4\"><item x=\"1\"/><item x=\"0\"/><item x=\"2\"/><item t=\"default\"/></items><autoSortScope><pivotArea dataOnly=\"0\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\"><x v=\"1\"/></reference></references></pivotArea></autoSortScope></pivotField>"
  )
  got <- xml_node(wb$pivotTables[1], "pivotTableDefinition", "pivotFields", "pivotField")[c(2, 10)]
  expect_equal(exp, got)

  expect_warning(
    wb$add_pivot_table(df, dims = "A3", rows = "cyl", cols = "gear",
                       data = c("vs", "am"), params = list(sort_row = 1, sort_col = -7)),
    "invalid sort position found"
  )

  expect_error(
    wb$add_pivot_table(df, dims = "A3", rows = "cyl", cols = "gear",
                       data = c("vs", "am"),
                       params = list(numfmt = c(numfmt = 10))),
    "length of numfmt and data does not match"
  )

  ### add sortType only to those pivot fields that are sorted

  ## sort by column and row
  df <- mtcars

  ## Create the workbook and the pivot table
  wb <- wb_workbook()$
    add_worksheet("Data")$
    add_data(x = df, start_col = 1, start_row = 2)

  df <- wb_data(wb)
  wb$add_pivot_table(
    df,
    dims = "A3",
    rows = c("cyl", "am"),
    cols = c("gear", "carb"),
    data = c("disp", "mpg"),
    params = list(sort_row = 1,
                  sort_col = -2)
  )

  exp <- c("", "ascending", "", "", "", "", "", "", "", "descending", "")
  got <- rbindlist(xml_attr(wb$pivotTables, "pivotTableDefinition", "pivotFields", "pivotField"))$sortType
  expect_equal(exp, got)

})

test_that("sort_item with pivot tables works", {

wb <- wb_workbook() %>% wb_add_worksheet() %>% wb_add_data(x = mtcars)

df <- wb_data(wb, sheet = 1)

expect_silent(
  wb_add_pivot_table(wb, df, dims = "A3",
                     filter = "am", rows = "cyl", cols = "gear", data = "disp",
                     params = list(sort_item = list(gear = c(3, 2, 1)))
  )
)

expect_warning(
  wb_add_pivot_table(wb, df, dims = "A3",
                     filter = "am", rows = "cyl", cols = "gear", data = "disp",
                     params = list(sort_item = list(gear = seq_len(4)))
  ),
  "Length of sort order for 'gear' does not match required length. Is 4, needs 3."
)

})

test_that("wbWorkbook print works", {

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_worksheet("Sheet 1 (1)")$
    add_worksheet("Sheet & NoSheet")

  exp <- c("A Workbook object.",
           " ",
           "Worksheets:",
           " Sheets: Sheet 1, Sheet 1 (1), Sheet & NoSheet ",
           " Write order: 1, 2, 3")
  got <- capture.output(wb)
  expect_equal(exp, got)

})

test_that("genBaseWorkbook() works", {

  # kinda superfluous. if it wouldn't work openxlsx2 would be broken

  exp <- c(
    "fileVersion", "fileSharing", "workbookPr", "alternateContent",
    "revisionPtr", "absPath", "workbookProtection", "bookViews",
    "sheets", "functionGroups", "externalReferences", "definedNames",
    "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
    "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects",
    "extLst"
  )
  expect_equal(
    names(genBaseWorkbook()),
    exp
  )

})

test_that("subsetting wb_data() works", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = head(esoph, 3))

  df1 <- wb_data(wb)

  exp <- data.frame(alcgp = rep("0-39g/day", 3), stringsAsFactors = FALSE)
  got <- df1[, 2, drop = FALSE]
  expect_equal(exp, got, ignore_attr = TRUE)

  exp <- rep("0-39g/day", 3)
  got <- unclass(df1[, 2])
  expect_equal(exp, got)

  exp <- structure(
    list(
      agegp = c("25-34", "25-34"),
      alcgp = c("0-39g/day", "0-39g/day")
    ),
    row.names = 2:3,
    dims = structure(
      list(
        A = c("A1", "A2", "A3"),
        B = c("B1", "B2", "B3")
      ),
      row.names = c(NA, 3L),
      class = "data.frame"),
    sheet = "Sheet 1"
  )
  got <- unclass(df1[1:2, 1:2])
  expect_equal(exp, got)

  expect_null(attributes(df1[c("agegp")]))

  exp <- list(
    names = "agegp",
    row.names = 2:4,
    class = c("wb_data", "data.frame"),
    dims = structure(
      list(A = c("A1", "A2", "A3", "A4")),
      row.names = c(NA, 4L),
      class = "data.frame"
    ),
    sheet = "Sheet 1"
  )
  got <- attributes(df1[c("agegp"), drop = FALSE])
  expect_equal(exp, got)

  exp <- list(
    names = c("alcgp", "tobgp", "ncases"),
    row.names = 2:4,
    class = c("wb_data", "data.frame"),
    dims = structure(
      list(
        B = c("B1", "B2", "B3", "B4"),
        C = c("C1", "C2", "C3", "C4"),
        D = c("D1", "D2", "D3", "D4")
      ),
      row.names = c(NA, 4L),
      class = "data.frame"
    ),
    sheet = "Sheet 1"
  )
  got <- attributes(df1[c("alcgp", "tobgp", "ncases")])
  expect_equal(exp, got)

  exp <- list(
    names = c("agegp", "alcgp", "tobgp", "ncases", "ncontrols"),
    row.names = 3:4,
    class = c("wb_data", "data.frame"),
    dims = structure(
      list(
        A = c("A1", "A3", "A4"),
        B = c("B1", "B3", "B4"),
        C = c("C1", "C3", "C4"),
        D = c("D1", "D3", "D4"),
        E = c("E1", "E3", "E4")
      ),
      row.names = c(1L, 3L, 4L),
      class = "data.frame"
    ),
    sheet = "Sheet 1"
  )
  got <- attributes(df1[-1, ])
  expect_equal(exp, got)

  exp <- list(
    names = c("agegp", "alcgp", "tobgp", "ncases", "ncontrols"),
    row.names = 2:3,
    class = c("wb_data", "data.frame"),
    dims = structure(
      list(
        A = c("A1", "A2", "A3"),
        B = c("B1", "B2", "B3"),
        C = c("C1", "C2", "C3"),
        D = c("D1", "D2", "D3"),
        E = c("E1", "E2", "E3")
      ),
      row.names = c(NA, 3L),
      class = "data.frame"
    ),
    sheet = "Sheet 1")
  got <- attributes(df1[-nrow(df1), ])
  expect_equal(exp, got)

})

test_that("adding mips section works", {

  # helper function to mock a mips section
  create_fake_mips <- function() {
    guid <- st_guid()
    lid  <- tolower(gsub("[{}]", "", guid))

    mips_xml <- sprintf(
      '<property fmtid="%s" pid="2" name="MSIP_Label_%s_Enabled">
      <vt:lpwstr>true</vt:lpwstr>
    </property>
    <property fmtid="%s" pid="3" name="MSIP_Label_%s_SetDate">
      <vt:lpwstr>2024-04-07T14:27:12Z</vt:lpwstr>
    </property>
    <property fmtid="%s" pid="4" name="MSIP_Label_%s_Method">
      <vt:lpwstr>Privileged</vt:lpwstr>
    </property>
    <property fmtid="%s" pid="5" name="MSIP_Label_%s_Name">
      <vt:lpwstr>General</vt:lpwstr>
    </property>
    <property fmtid="%s" pid="6" name="MSIP_Label_%s_SiteId">
      <vt:lpwstr>%s</vt:lpwstr>
    </property>
    <property fmtid="%s" pid="7" name="MSIP_Label_%s_ActionId">
      <vt:lpwstr>%s</vt:lpwstr>
    </property>
    <property fmtid="%s" pid="8" name="MSIP_Label_%s_ContentBits">
      <vt:lpwstr>0</vt:lpwstr>
    </property>',
      guid, lid,
      guid, lid,
      guid, lid,
      guid, lid,
      guid, lid, lid,
      guid, lid, lid,
      guid, lid
    )
    read_xml(mips_xml, pointer = FALSE)
  }

  fmips <- create_fake_mips()

  wb <- wb_workbook()$add_worksheet()$add_mips(xml = fmips)

  expect_message(wb_get_mips(wb, quiet = FALSE), "Found MIPS section: General")

  expect_equal(fmips, wb$get_mips())

  tmp <- temp_xlsx()

  wb$save(tmp)

  expect_equal(fmips, wb_load(tmp)$get_mips())

  wb <- wb_load(tmp)

  expect_message(wb$add_mips(xml = fmips), "File has duplicated custom section")
  expect_equal(fmips, wb$get_mips())


  op <- options("openxlsx2.mips_xml_string" = fmips)
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()$add_worksheet()$add_mips(xml = fmips)

  wb <- wb_workbook()$add_worksheet()$add_mips()
  expect_equal(fmips, wb$get_mips())


  wb <- wb_workbook() |>
    wb_add_worksheet() |>
    wb_set_properties(
      custom = list(
        Software    = "openxlsx2",
        Version     = 1.5,
        ReleaseDate = as.Date("2024-03-26"),
        CRAN        = TRUE,
        DEV         = FALSE
      )
    )

  wb <- wb_workbook()$add_worksheet()$
    set_properties(
      custom = list(int     = 1L)
    )
  expect_silent(wb$add_mips())

  exp <- "3"
  got <- rbindlist(xml_attr(wb$get_mips(single_xml = FALSE)[1], "property"))$pid
  expect_equal(exp, got)

})

test_that("handling mips in docMetadata works", {
  tmp <- temp_xlsx()
  xml <- '<clbl:labelList xmlns:clbl=\"http://schemas.microsoft.com/office/2020/mipLabelMetadata\"><clbl:label foo="bar"/></clbl:labelList>'
  wb <- wb_workbook() %>% wb_add_worksheet() %>% wb_add_mips(xml = xml)
  wb$docMetadata
  wb$save(tmp)
  rm(wb)

  wb <- wb_load(tmp)
  expect_equal(xml, wb$docMetadata)
})
