
test_that("add_data_table() writes over tables", {

  overwrite_table_error <- "Cannot overwrite existing table with another table"
  df1 <- data.frame("X" = 1:10)

  wb <- wb_add_worksheet(wb_workbook(), "Sheet1")

  ## table covers rows 4->10 and cols 4->8
  wb$add_data_table(sheet = 1, x = head(iris), start_col = 4, start_row = 4)

  ## should all run without error
  wb$add_data_table(sheet = 1, x = df1, start_col = 3, start_row = 2)
  wb$add_data_table(sheet = 1, x = df1, start_col = 9, start_row = 2)
  wb$add_data_table(sheet = 1, x = df1, start_col = 4, start_row = 11)
  wb$add_data_table(sheet = 1, x = df1, start_col = 5, start_row = 11)
  wb$add_data_table(sheet = 1, x = df1, start_col = 6, start_row = 11)
  wb$add_data_table(sheet = 1, x = df1, start_col = 7, start_row = 11)
  wb$add_data_table(sheet = 1, x = df1, start_col = 8, start_row = 11)
  wb$add_data_table(sheet = 1, x = head(iris, 2), start_col = 4, start_row = 1)

  ## Now error
  expect_error(wb$add_data_table(sheet = 1, x = df1, start_col = "H", start_row = 21), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = df1, start_col = 3, start_row = 12), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = df1, start_col = 9, start_row = 12), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = df1, start_col = "i", start_row = 12), regexp = overwrite_table_error)

  ## more errors
  expect_error(wb$add_data_table(sheet = 1, x = head(iris)), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), start_col = 4, start_row = 21), regexp = overwrite_table_error)

  ## should work
  wb$add_data_table(sheet = 1, x = head(iris), start_col = 4, start_row = 22)
  wb$add_data_table(sheet = 1, x = head(iris), start_col = 4, start_row = 40)

  ## more errors
  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), start_col = 4, start_row = 38), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), start_col = 4, start_row = 38, col_names = FALSE), regexp = overwrite_table_error)

  expect_error(wb$add_data_table(sheet = 1, x = head(iris), start_col = "H", start_row = 40), regexp = overwrite_table_error)
  wb$add_data_table(sheet = 1, x = head(iris), start_col = "I", start_row = 40)
  wb$add_data_table(sheet = 1, x = head(iris)[, 1:3], start_col = "A", start_row = 40)

  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), start_col = 4, start_row = 38, col_names = FALSE), regexp = overwrite_table_error)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris, 2), start_col = 1, start_row = 46, col_names = FALSE), regexp = overwrite_table_error)
})

test_that("zero row data table works", {

  wb <- wb_workbook()$
    add_worksheet()

  expect_warning(
    wb$add_data_table(x = data.frame(a = NA, b = NA)[0, ]),
    "Found data table with zero rows, adding one. Modify na with na.strings"
  )

  exp <- "A1:B2"
  got <- wb$tables$tab_ref
  expect_equal(exp, got)

})

test_that("write_data over tables", {
  overwrite_table_error <- "Cannot overwrite table headers. Avoid writing over the header row"
  df1 <- data.frame("X" = 1:10)

  wb <- wb_add_worksheet(wb_workbook(), "Sheet1")

  ## table covers rows 4->10 and cols 4->8
  wb$add_data_table(sheet = 1, x = head(iris), start_col = 4, start_row = 4)

  ## Anywhere on row 5 is fine
  for (i in 1:10) {
    wb$add_data(sheet = 1, x = head(iris), start_row = 5, start_col = i)
  }

  ## Anywhere on col i is fine
  for (i in 1:10) {
    wb$add_data(sheet = 1, x = head(iris), start_row = i, start_col = "i")
  }

  ## Now errors on headers
  expect_error(wb$add_data(sheet = 1, x = head(iris), start_col = 4, start_row = 4), regexp = overwrite_table_error)
  wb$add_data(sheet = 1, x = head(iris), start_col = 4, start_row = 5)
  wb$add_data(sheet = 1, x = head(iris)[1:3])
  wb$add_data(sheet = 1, x = head(iris, 2), start_col = 4)
  wb$add_data(sheet = 1, x = head(iris, 2), start_col = 4, col_names = FALSE)

  ## Example of how this should be used
  wb$add_data_table(sheet = 1, x = head(iris), start_col = 4, start_row = 30)
  wb$add_data(sheet = 1, x = head(iris), start_col = 4, start_row = 31, col_names = FALSE)

  wb$add_data_table(sheet = 1, x = head(iris), start_col = 10, start_row = 30)
  wb$add_data(sheet = 1, x = tail(iris), start_col = 10, start_row = 31, col_names = FALSE)

  wb$add_data_table(sheet = 1, x = head(iris)[, 1:3], start_col = 1, start_row = 30)
  wb$add_data(sheet = 1, x = tail(iris), start_col = 1, start_row = 31, col_names = FALSE)
})

test_that("Validate Table Names", {
  wb <- wb_add_worksheet(wb_workbook(), "Sheet 1")

  ## case
  expect_equal(wb_validate_table_name(wb, "test"), "test")
  expect_equal(wb_validate_table_name(wb, "TEST"), "test")
  expect_equal(wb_validate_table_name(wb, "Test"), "test")

  ## length
  expect_error(wb_validate_table_name(wb, paste(sample(LETTERS, size = 300, replace = TRUE), collapse = "")), regexp = "`table_name` must be less than 255 characters")

  ## look like cell ref
  expect_error(wb_validate_table_name(wb, "R1C2"), regexp = "`table_name` cannot be the same as a cell reference, such as R1C1", fixed = TRUE)
  expect_error(wb_validate_table_name(wb, "A1"), regexp = "`table_name` cannot be the same as a cell reference", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "R06821C9682"), regexp = "`table_name` cannot be the same as a cell reference, such as R1C1", fixed = TRUE)
  expect_error(wb_validate_table_name(wb, "ABD918751"), regexp = "`table_name` cannot be the same as a cell reference", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "A$100"), regexp = "`table_name` cannot contain spaces or the '$' character", fixed = TRUE)
  expect_error(wb_validate_table_name(wb, "A12$100"), regexp = "`table_name` cannot contain spaces or the '$' character", fixed = TRUE)

  tbl_nm <- "性別"
  expect_equal(wb_validate_table_name(wb, tbl_nm), tbl_nm)
})

test_that("Existing Table Names", {
  wb <- wb_add_worksheet(wb_workbook(), "Sheet 1")

  ## Existing names - case in-sensitive
  wb$add_data_table(sheet = 1, x = head(iris), table_name = "Table1")
  expect_error(wb_validate_table_name(wb, "Table1"), regexp = "`table_name = 'table1'` already exists", fixed = TRUE)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), table_name = "Table1", start_col = 10), regexp = "`table_name = 'table1'` already exists", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "TABLE1"), regexp = "`table_name = 'table1'` already exists", fixed = TRUE)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), table_name = "TABLE1", start_col = 20), regexp = "`table_name = 'table1'` already exists", fixed = TRUE)

  expect_error(wb_validate_table_name(wb, "table1"), regexp = "`table_name = 'table1'` already exists", fixed = TRUE)
  expect_error(wb$add_data_table(sheet = 1, x = head(iris), table_name = "table1", start_col = 30), regexp = "`table_name = 'table1'` already exists", fixed = TRUE)
})

test_that("custom table styles work", {

  # at the moment we have no interface to add custom table styles
  wb <- wb_workbook()$
    add_worksheet()

  # create dxf elements to be used in the table style
  tabCol1 <- create_dxfs_style(bg_fill = wb_color(theme = 7))
  tabCol2 <- create_dxfs_style(bg_fill = wb_color(theme = 5))
  tabBrd1 <- create_dxfs_style(border = TRUE)
  tabCol3 <- create_dxfs_style(bg_fill = wb_color(hex = "FFC00000"), font_color = wb_color("white"))

  # don't forget to assign them to the workbook
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

  expect_silent(wb$add_data_table(x = mtcars, table_style = "RedTableStyle"))
  wb$add_worksheet()
  expect_error(wb$add_data_table(x = mtcars, table_style = "RedTableStyle1"), "Invalid table style.")

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

  wb <- wb_workbook()$add_worksheet()$add_data_table(x = mtcars, with_filter = FALSE)
  wb$update_table(tabname = "Table1", dims = "A1:J4")

  got <- wb$tables$tab_ref
  expect_equal(got, "A1:J4")

})

test_that("writing table headers with factor variables works", {
  df <- data.frame(
    fct = structure(
      1:2,
      levels = c(
        "one",
        "two"
      ),
      class = "factor"),
    `1` = 1:2,
    check.names = FALSE
  )

  wb <- wb_workbook()$add_worksheet()$add_data_table(x = df)

  exp <- "<is><t>1</t></is>"
  got <- wb$worksheets[[1]]$sheet_data$cc$is[[2]]
  expect_equal(exp, got)
})

test_that("tables cannot have duplicated column names", {

  df1 <- data.frame(
    x = 1,
    x = 2,
    check.names = FALSE
  )

  df2 <- data.frame(
    x = 1,
    X = 2,
    check.names = FALSE
  )


  wb <- wb_workbook()

  expect_warning(wb$add_worksheet()$add_data_table(x = df1), "tables cannot have duplicated column names")
  expect_warning(wb$add_worksheet()$add_data_table(x = df2), "tables cannot have duplicated column names")


  exp <- c("<tableColumn id=\"1\" name=\"x\"/>", "<tableColumn id=\"2\" name=\"x2\"/>")
  got <- xml_node(wb$tables$tab_xml[[1]], "table", "tableColumns", "tableColumn")
  expect_equal(exp, got)

  exp <- c("<tableColumn id=\"1\" name=\"x\"/>", "<tableColumn id=\"2\" name=\"X2\"/>")
  got <- xml_node(wb$tables$tab_xml[[2]], "table", "tableColumns", "tableColumn")
  expect_equal(exp, got)

})

test_that("make sure that table id is unique", {

  tmp <- temp_xlsx()

  wb <- write_xlsx(x = list(head(mtcars), head(iris)), as_table = TRUE)

  wb$remove_tables(sheet = 1, table = "Table1", remove_data = FALSE)

  wb$save(tmp)
  wb <- wb_load(tmp)

  wb$add_worksheet()$add_data_table(x = head(Orange))

  exp <- data.frame(name = c("Table2", "Table3"), id = c("2", "3"), stringsAsFactors = FALSE)
  got <- rbindlist(xml_attr(wb$tables$tab_xml, "table"))[c("name", "id")]
  expect_equal(exp, got)

})

test_that("wb_get_named_regions, works with removed tables", {

  wb <- wb_workbook()$
    add_worksheet()$add_data_table(x = iris, table_name = "iris1")$
    add_worksheet()$add_data_table(x = iris, table_name = "iris2")$
    add_worksheet()$add_data_table(x = iris, table_name = "iris3")

  wb$remove_tables(table = "iris3")

  tabs <- wb$get_named_regions(tables = TRUE)

  expect_equal(2, nrow(tabs))
  expect_equal(paste0("iris", 1:2), tabs$name)

})

test_that("reading tables from file works", {
  tmp <- temp_xlsx()
  exp <- head(mtcars)
  write_xlsx(x = exp, file = tmp, as_table = TRUE)

  got <- wb_to_df(tmp, named_region = "Table1")
  expect_equal(exp, got, ignore_attr = TRUE)
})

test_that("remove_tables works", {

  df_h <- head(mtcars)
  df_t <- tail(mtcars)
  df_s <- mtcars[sample(seq_len(nrow(mtcars)), 6), ]

  dims_h <- wb_dims(x = df_h)
  dims_s <- wb_dims(x = df_s, from_dims = dims_h, below = 3)
  dims_t <- wb_dims(x = df_t, from_dims = dims_s, below = 3)


  wb <- wb_workbook()$add_worksheet()$
    add_data_table(x = df_h, dims = dims_h)$
    add_data(x = df_s, dims = dims_s)$
    add_data_table(x = df_t, dims = dims_t)

  expect_error(wb$remove_tables(table = "foo"), "does not exist")

  expect_silent(wb$remove_tables())

  expect_true(all(grepl("_deleted", wb$tables$tab_name)))

})
