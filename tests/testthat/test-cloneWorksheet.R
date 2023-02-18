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

test_that("copy cells", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)$
    add_fill(dims = "A1:F1", color = wb_color("yellow"))

  dat <- wb_data(wb, dims = "A1:D4", colNames = FALSE)

  # FIXME there is a bug with next_sheet() in clone_worksheets()
  wb$
    # 1:1 copy to M2
    clone_worksheet(old = 1, new = "Clone1")$
    copy_cells(data = dat, dims = "M2", as_value = FALSE, as_ref = FALSE, transpose = FALSE)$
    # 1:1 transposed copy to A20
    clone_worksheet(old = 1, new = "Clone2")$
    copy_cells(data = dat, dims = "A20", as_value = FALSE, as_ref = FALSE, transpose = TRUE)$
    # reference transposed copy to A20
    clone_worksheet(old = 1, new = "Clone3")$
    copy_cells(data = dat, dims = "A20", as_value = FALSE, as_ref = TRUE, transpose = TRUE)$
    # value copy to A20
    clone_worksheet(old = 1, new = "Clone4")$
    copy_cells(data = dat, dims = "A20", as_value = TRUE, as_ref = FALSE, transpose = FALSE)$
    # transposed value copy to A20
    clone_worksheet(old = 1, new = "Clone5")$
    copy_cells(data = dat, dims = "A20", as_value = TRUE, as_ref = FALSE, transpose = TRUE)

  got <- wb_data(wb, sheet = 2, dims = "M2:P5", colNames = FALSE)
  expect_equal(dat, got, ignore_attr = TRUE)

  got <- wb_data(wb, sheet = 3, dims = "A20:D23", colNames = FALSE)
  expect_equal(unlist(t(dat)), unlist(got), ignore_attr = TRUE)

  exp <- c("'Sheet 1'!A1", "'Sheet 1'!B1", "'Sheet 1'!C1", "'Sheet 1'!D1")
  got <- wb_data(wb, sheet = 4, dims = "A20:D23", colNames = FALSE, showFormula = TRUE)[[1]]
  expect_equal(exp, got)

  got <- wb_data(wb, sheet = 5, dims = "A20:D23", colNames = FALSE)
  expect_equal(dat, got, ignore_attr = TRUE)

  got <- wb_data(wb, sheet = 6, dims = "A20:D23", colNames = FALSE)
  expect_equal(unlist(t(dat)), unlist(got), ignore_attr = TRUE)

})
