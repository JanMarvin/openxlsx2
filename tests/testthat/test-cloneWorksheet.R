test_that("clone Worksheet with data", {
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_data("Sheet 1", 1)
  wb$clone_worksheet("Sheet 1", "Sheet 2")

  file_name <- system.file("extdata", "cloneWorksheetExample.xlsx", package = "openxlsx2")
  refwb <- wb_load(file = file_name)

  expect_equal(wb$get_sheet_names(), refwb$get_sheet_names())
  expect_equal(wb_get_order(wb), wb_get_order(refwb))
})


test_that("clone empty Worksheet", {
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$clone_worksheet("Sheet 1", "Sheet 2")

  file_name <- system.file("extdata", "cloneEmptyWorksheetExample.xlsx", package = "openxlsx2")
  refwb <- wb_load(file = file_name)

  expect_equal(wb$get_sheet_names(), refwb$get_sheet_names())
  expect_equal(wb_get_order(wb), wb_get_order(refwb))
})


test_that("clone Worksheet with table", {

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$add_data_table(sheet = "Sheet 1", x = iris, tableName = "iris")
  wb$add_data_table(sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)

  # FIXME a wild drawing2.xml appears in wb$Content_Types
  wb$clone_worksheet("Sheet 1", "Clone 1")

  old <- wb_validate_sheet(wb, "Sheet 1")
  new <- wb_validate_sheet(wb, "Clone 1")

  expect_equal(nrow(wb$tables), 4)
  expect_equal(nrow(wb$tables$tab_sheet == old), nrow(wb$tables$tab_sheet == new))
  relships <- rbindlist(xml_attr(unlist(wb$worksheets_rels), "Relationship"))
  relships$typ <- basename(relships$Type)
  relships$target <- basename(relships$Target)

  got <- relships[relships$typ == "table", c("Id", "typ", "target")]

  exp <- structure(list(
    Id = c("rId1", "rId2", "rId1", "rId2"),
    typ = c("table", "table", "table", "table"),
    target = c("table1.xml", "table2.xml", "table3.xml", "table4.xml")
  ),
  row.names = c(1L, 2L, 3L, 4L),
  class = "data.frame")

  expect_equal(got, exp)

})
