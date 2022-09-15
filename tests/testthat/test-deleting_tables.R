test_that("Deleting a Table Object", {
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_worksheet("Sheet 2")
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)


  # Get table ----

  expect_equal(length(wb_get_tables(wb, sheet = 1)), 2L)
  expect_equal(length(wb_get_tables(wb, sheet = "Sheet 1")), 2L)

  expect_equal(length(wb_get_tables(wb, sheet = 2)), 0)
  expect_equal(length(wb_get_tables(wb, sheet = "Sheet 2")), 0)

  expect_error(wb_get_tables(wb, sheet = 3), "Invalid sheet position")
  expect_error(wb_get_tables(wb, sheet = "Sheet 3"))

  expect_equal(wb_get_tables(wb, sheet = 1), c("iris", "mtcars"), ignore_attr = TRUE)
  expect_equal(wb_get_tables(wb, sheet = "Sheet 1"), c("iris", "mtcars"), ignore_attr = TRUE)

  expect_equal(attr(wb_get_tables(wb, sheet = 1), "refs"), c("A1:E151", "J1:T33"))
  expect_equal(attr(wb_get_tables(wb, sheet = "Sheet 1"), "refs"), c("A1:E151", "J1:T33"))

  expect_equal(nrow(wb$tables), 2L)

  ## Deleting a worksheet ----

  wb$remove_worksheet(1)
  expect_equal(nrow(wb$tables), 2L)
  expect_equal(length(wb_get_tables(wb, sheet = 1)), 0)

  expect_equal(wb$tables$tab_name, c("iris_openxlsx_deleted", "mtcars_openxlsx_deleted"))
  expect_equal(wb$tables$tab_sheet, c(0, 0))

  # wb$save(temp_xlsx())

  ## write same tables again ----

  wb$add_data_table(sheet = 1, x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)

  expect_equal(wb$tables$tab_name, c("iris_openxlsx_deleted", "mtcars_openxlsx_deleted", "iris", "mtcars"))
  expect_equal(wb$tables$tab_sheet, c(0, 0, 1, 1))

  expect_equal(length(wb_get_tables(wb, sheet = 1)), 2L)
  expect_equal(length(wb_get_tables(wb, sheet = "Sheet 2")), 2L)

  expect_error(wb_get_tables(wb, sheet = 2))
  expect_error(wb_get_tables(wb, sheet = "Sheet 1"))

  expect_equal(wb_get_tables(wb, sheet = 1), c("iris", "mtcars"), ignore_attr = TRUE)
  expect_equal(wb_get_tables(wb, sheet = "Sheet 2"), c("iris", "mtcars"), ignore_attr = TRUE)

  expect_equal(attr(wb_get_tables(wb, sheet = 1), "refs"), c("A1:E151", "J1:T33"))
  expect_equal(attr(wb_get_tables(wb, sheet = "Sheet 2"), "refs"), c("A1:E151", "J1:T33"))

  expect_equal(nrow(wb$tables), 4L)


  ## wb_remove_tables ----

  ## remove iris and re-write it
  wb$remove_tables(sheet = 1, table = "iris")

  expect_equal(nrow(wb$tables), 4L)
  expect_equal(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId2\"/>", ignore_attr = TRUE)
  expect_equal(attr(wb$worksheets[[1]]$tableParts, "tableName"), "mtcars")

  expect_equal(wb$tables$tab_name, c(
    "iris_openxlsx_deleted",
    "mtcars_openxlsx_deleted",
    "iris_openxlsx_deleted",
    "mtcars"
  ))

  ## wb_remove_tables clears table object and all data
  wb$add_data_table(sheet = 1, x = iris, tableName = "iris", startCol = 1)
  temp <- temp_xlsx()
  wb_save(wb, temp)
  expect_equal(wb$worksheets[[1]]$tableParts, c("<tablePart r:id=\"rId2\"/>", "<tablePart r:id=\"rId3\"/>"), ignore_attr = TRUE)
  expect_equal(attr(wb$worksheets[[1]]$tableParts, "tableName"), c("mtcars", "iris"))


  wb$remove_tables(sheet = 1, table = "iris")

  expect_equal(nrow(wb$tables), 5L)
  expect_equal(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId2\"/>", ignore_attr = TRUE)
  expect_equal(attr(wb$worksheets[[1]]$tableParts, "tableName"), "mtcars")

  expect_equal(wb$tables$tab_name, c(
    "iris_openxlsx_deleted",
    "mtcars_openxlsx_deleted",
    "iris_openxlsx_deleted",
    "mtcars",
    "iris_openxlsx_deleted"
  ))


  expect_equal(wb_get_tables(wb, sheet = 1), "mtcars", ignore_attr = TRUE)
  file.remove(temp)
})

test_that("Save and load Table Deletion", {
  temp_file <- temp_xlsx()

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_worksheet("Sheet 2")
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)


  ###################################################################################
  ## Deleting a worksheet

  wb$remove_worksheet(1)
  expect_equal(nrow(wb$tables), 2L)
  expect_equal(length(wb_get_tables(wb, sheet = 1)), 0)

  expect_equal(wb$tables$tab_name, c("iris_openxlsx_deleted", "mtcars_openxlsx_deleted"))
  expect_equal(wb$tables$tab_sheet, c(0, 0))


  ## both table were written to sheet 1 and are expected to not exist after load
  wb_save(wb, temp_file)
  wb <- wb_load(file = temp_file)
  expect_null(wb$tables)
  file.remove(temp_file)




  ###################################################################################
  ## Deleting a table

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_worksheet("Sheet 2")
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)

  ## remove iris and re-write it
  wb$remove_tables(sheet = 1, table = "iris")
  expect_equal(wb$tables$tab_name, c("iris_openxlsx_deleted", "mtcars"))

  temp_file <- temp_xlsx()
  wb$save(temp_file)
  wb <- wb_load(file = temp_file)

  expect_equal(nrow(wb$tables), 1L)
  expect_equal(wb$tables$tab_name, "mtcars")

  expect_equal(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId2\"/>", ignore_attr = TRUE) ## rId reset
  expect_equal(unname(attr(wb$worksheets[[1]]$tableParts, "tableName")), "mtcars")
  file.remove(temp_file)



  ## now delete the other table
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_worksheet("Sheet 2")
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
  wb$add_data_table(sheet = 2, x = mtcars, tableName = "mtcars2", startCol = 3)

  wb$remove_tables(sheet = 1, table = "iris")
  wb$remove_tables(sheet = 1, table = "mtcars")
  expect_equal(wb$tables$tab_name, c("iris_openxlsx_deleted", "mtcars_openxlsx_deleted", "mtcars2"))

  temp_file <- temp_xlsx()
  wb_save(wb, temp_file)
  wb <- wb_load(file = temp_file)

  expect_equal(nrow(wb$tables), 1L)
  expect_equal(wb$tables$tab_name, "mtcars2")
  expect_length(wb$worksheets[[1]]$tableParts, 0)
  expect_equal(wb$worksheets[[2]]$tableParts, "<tablePart r:id=\"rId1\"/>", ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[2]]$tableParts, "tableName")), "mtcars2")
  unlink(temp_file)


  ## write tables back in
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)

  expect_equal(nrow(wb$tables), 3L)
  expect_equal(wb$tables$tab_name, c("mtcars2", "iris", "mtcars"))

  expect_length(wb$worksheets[[1]]$tableParts, 2)
  expect_equal(wb$worksheets[[1]]$tableParts, c("<tablePart r:id=\"rId1\"/>", "<tablePart r:id=\"rId2\"/>"), ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[1]]$tableParts, "tableName")), c("iris", "mtcars"))

  expect_length(wb$worksheets[[2]]$tableParts, 1)
  expect_equal(wb$worksheets[[2]]$tableParts, c("<tablePart r:id=\"rId1\"/>"), ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[2]]$tableParts, "tableName")), "mtcars2")

  wb_save(wb, temp_file)


  ## Ids should get reset after load
  wb <- wb_load(file = temp_file)

  expect_equal(nrow(wb$tables), 3L)
  expect_equal(wb$tables$tab_name, c("iris", "mtcars", "mtcars2"))

  expect_length(wb$worksheets[[1]]$tableParts, 2)
  expect_equal(wb$worksheets[[1]]$tableParts, c("<tablePart r:id=\"rId1\"/>", "<tablePart r:id=\"rId2\"/>"), ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[1]]$tableParts, "tableName")), c("iris", "mtcars"))

  expect_length(wb$worksheets[[2]]$tableParts, 1)
  expect_equal(wb$worksheets[[2]]$tableParts, c("<tablePart r:id=\"rId1\"/>"), ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[2]]$tableParts, "tableName")), "mtcars2")

  unlink(temp_file)
})
