
test_that("write_datatable over tables", {

  overwrite_table_error <- "Cannot overwrite existing table with another table"
  df1 <- data.frame("X" = 1:10)

  wb <- wb_add_worksheet(wb_workbook(), "Sheet1")

  ## table covers rows 4->10 and cols 4->8
  wb$add_data_table(sheet = 1, x = head(iris), startCol = 4, startRow = 4)

  ## should all run without error
  wb$add_data_table(sheet = 1, x = df1, startCol = 3, startRow = 2)
  wb$add_data_table(sheet = 1, x = df1, startCol = 9, startRow = 2)
  wb$add_data_table(sheet = 1, x = df1, startCol = 4, startRow = 11)
  wb$add_data_table(sheet = 1, x = df1, startCol = 5, startRow = 11)
  wb$add_data_table(sheet = 1, x = df1, startCol = 6, startRow = 11)
  wb$add_data_table(sheet = 1, x = df1, startCol = 7, startRow = 11)
  wb$add_data_table(sheet = 1, x = df1, startCol = 8, startRow = 11)
  wb$add_data_table(sheet = 1, x = head(iris, 2), startCol = 4, startRow = 1)



  ## Now error
  expect_error(wb$add_data_table(sheet = 1, x = df1, startCol = "H", startRow = 21), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = df1, startCol = 3, startRow = 12), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = df1, startCol = 9, startRow = 12), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = df1, startCol = "i", startRow = 12), regexp = overwrite_table_error)


  ## more errors
  expect_error(wb$add_data_table(sheet = 1, x = head(iris)), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), startCol = 4, startRow = 21), regexp = overwrite_table_error)

  ## should work
  wb$add_data_table(sheet = 1, x = head(iris), startCol = 4, startRow = 22)
  wb$add_data_table(sheet = 1, x = head(iris), startCol = 4, startRow = 40)


  ## more errors
  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), startCol = 4, startRow = 38), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), startCol = 4, startRow = 38, colNames = FALSE), regexp = overwrite_table_error)

  expect_error(wb$add_data_table(sheet = 1, x = head(iris), startCol = "H", startRow = 40), regexp = overwrite_table_error)
  wb$add_data_table(sheet = 1, x = head(iris), startCol = "I", startRow = 40)
  wb$add_data_table(sheet = 1, x = head(iris)[, 1:3], startCol = "A", startRow = 40)

  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), startCol = 4, startRow = 38, colNames = FALSE), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), startCol = 1, startRow = 46, colNames = FALSE), regexp = overwrite_table_error)
})





test_that("write_data over tables", {
  overwrite_table_error <- "Cannot overwrite table headers. Avoid writing over the header row"
  df1 <- data.frame("X" = 1:10)

  wb <- wb_add_worksheet(wb_workbook(), "Sheet1")

  ## table covers rows 4->10 and cols 4->8
  wb$add_data_table(sheet = 1, x = head(iris), startCol = 4, startRow = 4)

  ## Anywhere on row 5 is fine
  for (i in 1:10) {
    wb$add_data(sheet = 1, x = head(iris), startRow = 5, startCol = i)
  }

  ## Anywhere on col i is fine
  for (i in 1:10) {
    wb$add_data(sheet = 1, x = head(iris), startRow = i, startCol = "i")
  }



  ## Now errors on headers
  expect_error(wb$add_data(sheet = 1, x = head(iris), startCol = 4, startRow = 4), regexp = overwrite_table_error)
  wb$add_data(sheet = 1, x = head(iris), startCol = 4, startRow = 5)
  wb$add_data(sheet = 1, x = head(iris)[1:3])
  wb$add_data(sheet = 1, x = head(iris, 2), startCol = 4)
  wb$add_data(sheet = 1, x = head(iris, 2), startCol = 4, colNames = FALSE)


  ## Example of how this should be used
  wb$add_data_table(sheet = 1, x = head(iris), startCol = 4, startRow = 30)
  wb$add_data(sheet = 1, x = head(iris), startCol = 4, startRow = 31, colNames = FALSE)

  wb$add_data_table(sheet = 1, x = head(iris), startCol = 10, startRow = 30)
  wb$add_data(sheet = 1, x = tail(iris), startCol = 10, startRow = 31, colNames = FALSE)

  wb$add_data_table(sheet = 1, x = head(iris)[, 1:3], startCol = 1, startRow = 30)
  wb$add_data(sheet = 1, x = tail(iris), startCol = 1, startRow = 31, colNames = FALSE)
})

test_that("Validate Table Names", {
  wb <- wb_add_worksheet(wb_workbook(), "Sheet 1")

  ## case
  expect_equal(wb_validate_table_name(wb, "test"), "test")
  expect_equal(wb_validate_table_name(wb, "TEST"), "test")
  expect_equal(wb_validate_table_name(wb, "Test"), "test")

  ## length
  expect_error(wb_validate_table_name(wb, paste(sample(LETTERS, size = 300, replace = TRUE), collapse = "")), regexp = "tableName must be less than 255 characters")

  ## look like cell ref
  expect_error(wb_validate_table_name(wb, "R1C2"), regexp = "tableName cannot be the same as a cell reference, such as R1C1", fixed = TRUE)
  expect_error(wb_validate_table_name(wb, "A1"), regexp = "tableName cannot be the same as a cell reference", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "R06821C9682"), regexp = "tableName cannot be the same as a cell reference, such as R1C1", fixed = TRUE)
  expect_error(wb_validate_table_name(wb, "ABD918751"), regexp = "tableName cannot be the same as a cell reference", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "A$100"), regexp = "'$' character cannot exist in a tableName", fixed = TRUE)
  expect_error(wb_validate_table_name(wb, "A12$100"), regexp = "'$' character cannot exist in a tableName", fixed = TRUE)

  tbl_nm <- "性別"
  expect_equal(wb_validate_table_name(wb, tbl_nm), tbl_nm)
})

test_that("Existing Table Names", {
  wb <- wb_add_worksheet(wb_workbook(), "Sheet 1")

  ## Existing names - case in-sensitive
  wb$add_data_table(sheet = 1, x = head(iris), tableName = "Table1")
  expect_error(wb_validate_table_name(wb, "Table1"), regexp = "table with name 'table1' already exists", fixed = TRUE)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), tableName = "Table1", startCol = 10), regexp = "table with name 'table1' already exists", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "TABLE1"), regexp = "table with name 'table1' already exists", fixed = TRUE)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), tableName = "TABLE1", startCol = 20), regexp = "table with name 'table1' already exists", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "table1"), regexp = "table with name 'table1' already exists", fixed = TRUE)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), tableName = "table1", startCol = 30), regexp = "table with name 'table1' already exists", fixed = TRUE)
})

test_that("custom table styles work", {

  # at the moment we have no interface to add custom table styles
  wb <- wb_workbook() %>%
    wb_add_worksheet()

  # create dxf elements to be used in the table style
  tabCol1 <- create_dxfs_style(bgFill = wb_color(theme = 7))
  tabCol2 <- create_dxfs_style(bgFill = wb_color(theme = 5))
  tabBrd1 <- create_dxfs_style(border = TRUE)
  tabCol3 <- create_dxfs_style(bgFill = wb_color(hex = "FFC00000"), font_color = wb_color("white"))

  # dont forget to assign them to the workbook
  wb$add_style(tabCol1)
  wb$add_style(tabCol2)
  wb$add_style(tabBrd1)
  wb$add_style(tabCol3)

  # tweak a working style with 4 elements
  tab_xml <- sprintf(
    "
     <tableStyle name=\"RedTableStyle\" pivot=\"0\" count=\"%s\" xr9:uid=\"{91A57EDA-14C5-4643-B7E3-C78161B6BBA4}\">
       <tableStyleElement type=\"wholeTable\" dxfId=\"%s\"/>
       <tableStyleElement type=\"headerRow\" dxfId=\"%s\"/>
       <tableStyleElement type=\"firstRowStripe\" dxfId=\"%s\"/>
       <tableStyleElement type=\"secondColumnStripe\" dxfId=\"%s\"/>
     </tableStyle>
    ",
    length(c(tabCol1, tabCol2, tabCol3, tabBrd1)),
    wb$styles_mgr$get_dxf_id("tabBrd1"),
    wb$styles_mgr$get_dxf_id("tabCol3"),
    wb$styles_mgr$get_dxf_id("tabCol1"),
    wb$styles_mgr$get_dxf_id("tabCol2")
  )
  wb$add_style(tab_xml)

  expect_silent(wb$add_data_table(x = mtcars, tableStyle = "RedTableStyle"))
  wb$add_worksheet()
  expect_error(wb$add_data_table(x = mtcars, tableStyle = "RedTableStyle1"), "Invalid table style.")

})

test_that("updating table works", {

  wb <- wb_workbook()
  wb$add_worksheet()$add_data_table(x = mtcars)
  wb$add_worksheet()$add_data_table(x = mtcars[1:2])

  wb_to_df(wb, named_region = "Table2")

  wb$add_data(dims = "C1", x = mtcars[-1:-2], name = "test")

  wb <- wb_update_table(wb, "Table2", sheet = 2, dims = "A1:J4")

  exp <- mtcars[1:3, 1:10]
  got <- wb_to_df(wb, named_region = "Table2")
  expect_equal(exp, got, ignore_attr = TRUE)

})
