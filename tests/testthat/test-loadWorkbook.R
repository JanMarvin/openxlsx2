
test_that("loadWorkbook() works", {
  wb <- loadWorkbook(system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))

  expect_identical(unname(attr(wb$tables, "tableName")), c("Table2", "Table3"))
  expect_identical(names(attr(wb$tables, "tableName")), c("A1:E51", "A1:K30"))
  expect_identical(attr(wb$tables, "sheet"), c(1L, 3L))

  expect_identical(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId3\"/>", ignore_attr = TRUE)
  expect_identical(unname(attr(wb$worksheets[[1]]$tableParts, "tableName")), "Table2")
  expect_identical(names(attr(wb$worksheets[[1]]$tableParts, "tableName")), "A1:E51")

  expect_identical(wb$worksheets[[3]]$tableParts, "<tablePart r:id=\"rId4\"/>", ignore_attr = TRUE)
  expect_identical(unname(attr(wb$worksheets[[3]]$tableParts, "tableName")), "Table3")
  expect_identical(names(attr(wb$worksheets[[3]]$tableParts, "tableName")), "A1:K30")


  ## now remove a table
  expect_identical(unname(getTables(wb, 1)), "Table2", ignore_attr = TRUE)
  expect_identical(unname(getTables(wb, 3)), "Table3", ignore_attr = TRUE)

  removeTable(wb, sheet = 1, table = "Table2")

  expect_identical(getTables(wb, sheet = 1), character(), ignore_attr = TRUE)
  expect_identical(length(wb$worksheets[[1]]$tableParts), 0L)
  expect_identical(wb$worksheets[[1]]$tableParts, character(), ignore_attr = TRUE)

  expect_identical(wb$worksheets[[3]]$tableParts, "<tablePart r:id=\"rId4\"/>", ignore_attr = TRUE)
  expect_identical(unname(attr(wb$worksheets[[3]]$tableParts, "tableName")), "Table3")
  expect_identical(names(attr(wb$worksheets[[3]]$tableParts, "tableName")), "A1:K30")

  expect_error(removeTable(wb, sheet = 1, table = "Table2"), regexp = "table 'Table2' does not exist")
})
