
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
