

test_that("Tables loaded correctly", {
  wb <- wb_load(system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))

  expect_equal(wb$tables$tab_name, c("Table2", "Table3"))
  expect_equal(wb$tables$tab_ref, c("A1:E51", "A1:K30"))
  expect_equal(wb$tables$tab_sheet, c(1, 3))

  expect_equal(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId3\"/>", ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[1]]$tableParts, "tableName")), "Table2")
  # expect_equal(names(attr(wb$worksheets[[1]]$tableParts, "tableName")), "A1:E51")

  expect_equal(wb$worksheets[[3]]$tableParts, "<tablePart r:id=\"rId2\"/>", ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[3]]$tableParts, "tableName")), "Table3")
  # expect_equal(names(attr(wb$worksheets[[3]]$tableParts, "tableName")), "A1:K30")


  ## now remove a table
  expect_equal(unname(wb_get_tables(wb, 1)), "Table2", ignore_attr = TRUE)
  expect_equal(unname(wb_get_tables(wb, 3)), "Table3", ignore_attr = TRUE)

  wb$remove_tables(sheet = 1, table = "Table2")

  expect_equal(wb_get_tables(wb, sheet = 1), character(), ignore_attr = TRUE)
  expect_equal(length(wb$worksheets[[1]]$tableParts), 0)
  expect_equal(wb$worksheets[[1]]$tableParts, character(), ignore_attr = TRUE)

  expect_equal(wb$worksheets[[3]]$tableParts, "<tablePart r:id=\"rId2\"/>", ignore_attr = TRUE)
  expect_equal(unname(attr(wb$worksheets[[3]]$tableParts, "tableName")), "Table3")
  # expect_equal(names(attr(wb$worksheets[[3]]$tableParts, "tableName")), "A1:K30")


  expect_error(wb_remove_tables(wb, sheet = 1, table = "Table2"), regexp = "table 'Table2' does not exist")
})
