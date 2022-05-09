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
