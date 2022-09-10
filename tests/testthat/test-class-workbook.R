test_that("Workbook class", {
  expect_null(assert_workbook(wb_workbook()))
})


test_that("wb_set_col_widths", {
# TODO use wb$wb_set_col_widths()

  wb <- wbWorkbook$new()
  wb$add_worksheet("test")
  wb$add_data("test", mtcars)

  # set column width to 12
  expect_silent(wb$set_col_widths("test", widths = 12L, cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"12\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # wrong sheet
  expect_error(wb$set_col_widths("test2", widths = 12L, cols = seq_along(mtcars)))

  # reset the column with, we do not provide an option ot remove the column entry
  expect_silent(wb$set_col_widths("test", cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"8.43\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # create column width for column 25
  expect_silent(wb$set_col_widths("test", cols = "Y", widths = 22))
  expect_equal(
    c("<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"8.43\"/>",
      "<col min=\"12\" max=\"24\" width=\"8.43\"/>",
      "<col min=\"25\" max=\"25\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"22\"/>"),
    wb$worksheets[[1]]$cols_attr
  )

  # a few more errors
  expect_error(wb$set_col_widths("test", cols = "Y", width = 1:2))
  expect_error(wb$set_col_widths("test", cols = "Y", hidden = 1:2))
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
    add_data_validation(col = 1:3, rows = 2:151, type = "whole",
                        operator = "between", value = c(1, 9)
    )$
    # text width 7-9 is fine
    add_data_validation(col = 5, rows = 2:151, type = "textLength",
                        operator = "between", value = c(7, 9)
    )$
    ## Date and Time cell validation
    add_worksheet("Sheet 2")$
    add_data_table(x = df)$
    # date >= 2016-01-01 is fine
    add_data_validation(col = 1, rows = 2:12, type = "date",
                        operator = "greaterThanOrEqual", value = as.Date("2016-01-01")
    )$
    # a few timestamps are fine
    add_data_validation(col = 2, rows = 2:12, type = "time",
                        operator = "between", value = df$t[c(4, 8)]
    )$
    ## validate list: validate inputs on one sheet with another
    add_worksheet("Sheet 3")$
    add_data_table(x = iris[1:30, ])$
    add_worksheet("Sheet 4")$
    add_data(x = sample(iris$Sepal.Length, 10))$
    add_data_validation("Sheet 3", col = 1, rows = 2:31, type = "list",
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
    "<x14:dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\"><x14:formula1><xm:f>'Sheet 4'!$A$1:$A$10</xm:f></x14:formula1><xm:sqref>A2:A31</xm:sqref></x14:dataValidation>"
  )
  got <- xml_node(wb$worksheets[[3]]$extLst, "ext", "x14:dataValidations", "x14:dataValidation")
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
    xml_node(wb$worksheets[[3]]$extLst, "ext", "x14:dataValidations", "x14:dataValidation"),
    xml_node(wb2$worksheets[[3]]$extLst, "ext", "x14:dataValidations", "x14:dataValidation")
  )

  wb2$add_data_validation("Sheet 3", col = 2, rows = 2:31, type = "list",
                          value = "'Sheet 4'!$A$1:$A$10")

  exp <- c(
    "<x14:dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\"><x14:formula1><xm:f>'Sheet 4'!$A$1:$A$10</xm:f></x14:formula1><xm:sqref>A2:A31</xm:sqref></x14:dataValidation>",
    "<x14:dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\"><x14:formula1><xm:f>'Sheet 4'!$A$1:$A$10</xm:f></x14:formula1><xm:sqref>B2:B31</xm:sqref></x14:dataValidation>"
  )
  got <- xml_node(wb2$worksheets[[3]]$extLst, "ext", "x14:dataValidations", "x14:dataValidation")
  expect_equal(exp, got)

  ### tests if conditions

  # test col2int
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data_table(x = head(iris))$
    # whole numbers are fine
    add_data_validation(col = "A", rows = 2:151, type = "whole",
                        operator = "between", value = c(1, 9)
    )

  exp <- "<dataValidation type=\"whole\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A151\"><formula1>1</formula1><formula2>9</formula2></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)


  # to many values
  expect_error(
    wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data_table(x = head(iris))$
    add_data_validation(col = "A", rows = 2:151, type = "whole",
                        operator = "between", value = c(1, 9, 19)
    ),
    "length <= 2"
  )

  # wrong type
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      add_data_validation(col = "A", rows = 2:151, type = "even",
                          operator = "between", value = c(1, 9)
      ),
    "Invalid 'type' argument!"
  )

  # wrong operator
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      add_data_validation(col = "A", rows = 2:151, type = "whole",
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
      add_data_validation(col = 1, rows = 2:12, type = "date",
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
      add_data_validation(col = 1, rows = 2:12, type = "time",
                          operator = "greaterThanOrEqual", value = 7
      ),
    "If type == 'time' value argument must be a POSIXct or POSIXlt vector."
  )


  # some more options
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = c(-1:1), colNames = FALSE)$
    # whole numbers are fine
    add_data_validation(col = 1, rows = 1:3, type = "whole",
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
    add_data_validation(col = 1, rows = 1:3, type = "custom", value = "A1=B1")

  exp <- "<dataValidation type=\"custom\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1:A3\"><formula1>A1=B1</formula1></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)

})


test_that("clone worksheet", {

  ## Dummy tests - not sure how to test these from R ##

  # # clone chartsheet ----------------------------------------------------
  fl <- system.file("extdata", "mtcars_chart.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  # wb$get_sheet_names() # chartsheet has no named name?
  expect_silent(wb$clone_worksheet(1, "Clone 1"))
  expect_true(inherits(wb$worksheets[[5]], "wbChartSheet"))
  # wb$open()

  # clone pivot table and drawing -----------------------------------------
  fl <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
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
  fl <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet("testing", "Clone1"))

  expect_false(identical(wb$worksheets_rels[2], wb$worksheets_rels[5]))
  # wb$open()

  # clone sheet with table ------------------------------------------------
  fl <- system.file("extdata", "tableStyles.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet(1, "clone"))

  expect_false(identical(wb$tables$tab_xml[1], wb$tables$tab_xml[2]))
  # wb$open()

  # clone sheet with chart ------------------------------------------------
  fl <- system.file("extdata", "mtcars_chart.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  wb$clone_worksheet(2, "Clone 1")

  expect_true(grepl("test", wb$charts$chart[2]))
  expect_true(grepl("'Clone 1'", wb$charts$chart[3]))
  # wb$open()

  # clone slicer ----------------------------------------------------------
  fl <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  expect_warning(wb$clone_worksheet("IrisSample", "Clone1"),
                 "Cloning slicers is not yet supported. It will not appear on the sheet.")
  # wb$open()

})
