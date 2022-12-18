test_that("group columns", {

  # Grouping then setting widths updates hidden
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")

  wb$createCols("Sheet 1", 3)

  wb$group_cols("Sheet 1", 2:3)
  wb$set_col_widths("Sheet 1", 2, widths = "18", hidden = FALSE)

  exp <- c(
    "<col min=\"1\" max=\"1\" width=\"8.43\"/>",
    "<col min=\"2\" max=\"2\" bestFit=\"1\" collapsed=\"0\" customWidth=\"1\" hidden=\"false\" outlineLevel=\"1\" width=\"18.711\"/>",
    "<col min=\"3\" max=\"3\" collapsed=\"0\" width=\"8.43\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)

  # Setting column widths then grouping
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")

  wb$createCols("Sheet 1", 3)

  wb$
    set_col_widths("Sheet 1", 2:3, widths = "18", hidden = FALSE)$
    group_cols("Sheet 1", 1:2, collapsed = TRUE)

  exp <- c(
    "<col min=\"1\" max=\"1\" collapsed=\"1\" hidden=\"1\" outlineLevel=\"1\" width=\"8.43\"/>",
    "<col min=\"2\" max=\"2\" bestFit=\"1\" collapsed=\"1\" customWidth=\"1\" hidden=\"false\" width=\"18.711\"/>",
    "<col min=\"3\" max=\"3\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"18.711\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)
})

test_that("group rows", {

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = cbind(rep(NA, 4)), na.strings = NULL, colNames = FALSE)$
    group_rows("Sheet 1", 1:4)

  got <- wb$worksheets[[1]]$sheet_data$row_attr$outlineLevel
  expect_equal(c("1", "1", "1", ""), got)

  got <- wb$worksheets[[1]]$sheet_data$row_attr$r
  expect_equal(c("1", "2", "3", "4"), got)

})

test_that("group rows 2", {

  wb <- wb_workbook()$
    add_worksheet('Sheet1')$add_data(x = iris)$
    add_worksheet('Sheet2')$add_data(x = iris)

  ## create list of groups
  # lines used for grouping (here: species)
  grp <- list(
    seq(2, 51),
    seq(52, 101),
    seq(102, 151)
  )

  # assign group levels
  names(grp) <- c("1", "0", "1")
  wb$group_rows("Sheet1", rows = grp)

  # different grouping
  names(grp) <- c("1", "2", "3")
  wb$group_rows("Sheet2", rows = grp)

  exp <- c("", "1", "0")
  got <- unique(wb$worksheets[[1]]$sheet_data$row_attr$outlineLevel)
  expect_equal(exp, got)

  exp <- c("", "1", "2", "3")
  got <- unique(wb$worksheets[[2]]$sheet_data$row_attr$outlineLevel)
  expect_equal(exp, got)
})

test_that("ungroup columns", {

  # OutlineLevelCol is removed from SheetFormatPr when no
  # column groupings left
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$createCols("Sheet 1", 3)
  wb$set_col_widths("Sheet 1", 2, widths = "18", hidden = FALSE)
  wb$group_cols("Sheet 1", 1:3)

  wb$ungroup_cols("Sheet 1", 1:3)

  exp <- c(
    "<col min=\"1\" max=\"1\" width=\"8.43\"/>",
    "<col min=\"2\" max=\"2\" bestFit=\"1\" customWidth=\"1\" width=\"18.711\"/>",
    "<col min=\"3\" max=\"3\" width=\"8.43\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)

})

test_that("ungroup rows", {

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = cbind(rep(NA, 4)), na.strings = NULL, colNames = FALSE)$
    group_rows("Sheet 1", 1:4)$
    ungroup_rows("Sheet 1", 1:4)

  got <- wb$worksheets[[1]]$sheet_data$row_attr$outlineLevel
  expect_true(all(got == ""))

})
