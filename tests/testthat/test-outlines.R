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

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = cbind(rep(NA, 4)), na.strings = NULL, colNames = FALSE)$
    group_rows("Sheet 1", 1:4, collapsed = TRUE)

  got <- wb$worksheets[[1]]$sheet_data$row_attr$collapsed
  expect_equal(c("1", "1", "1", "1"), got)

  got <- wb$worksheets[[1]]$sheet_data$row_attr$hidden
  expect_equal(c("1", "1", "1", ""), got)

})

test_that("grouping levels", {

  # create matrix
  t1 <- AirPassengers
  t2 <- do.call(cbind, split(t1, cycle(t1)))
  dimnames(t2) <- dimnames(.preformat.ts(t1))

  wb <- wb_workbook()
  wb$add_worksheet("AirPass")
  wb$add_data("AirPass", t2, rowNames = TRUE)

  # lines used for grouping (here: species)
  grp_rows <- list(
    "1" = seq(2, 3),
    "2" = seq(4, 8),
    "3" = seq(9, 13)
  )

  # lines used for grouping (here: quarter)
  grp_cols <- list(
    "1" = seq(2, 4),
    "2" = seq(5, 7),
    "3" = seq(8, 10),
    "4" = seq(11, 13)
  )

  wb <- wb_workbook()
  wb$add_worksheet("AirPass")
  wb$add_data("AirPass", t2, rowNames = TRUE)

  wb$createCols("AirPass", 13)

  wb$group_cols("AirPass", cols = grp_cols)
  wb$group_rows("AirPass", rows = grp_rows)

  exp <- c("", "1", "2", "3")
  got <- unique(wb$worksheets[[1]]$sheet_data$row_attr$outlineLevel)
  expect_equal(exp, got)

  exp <- c(
    "<col min=\"1\" max=\"1\" width=\"8.43\"/>", "<col min=\"2\" max=\"3\" collapsed=\"0\" hidden=\"0\" outlineLevel=\"1\" width=\"8.43\"/>",
    "<col min=\"4\" max=\"4\" collapsed=\"0\" hidden=\"0\" width=\"8.43\"/>",
    "<col min=\"5\" max=\"6\" collapsed=\"0\" hidden=\"0\" outlineLevel=\"2\" width=\"8.43\"/>",
    "<col min=\"7\" max=\"7\" collapsed=\"0\" hidden=\"0\" width=\"8.43\"/>",
    "<col min=\"8\" max=\"9\" collapsed=\"0\" hidden=\"0\" outlineLevel=\"3\" width=\"8.43\"/>",
    "<col min=\"10\" max=\"10\" collapsed=\"0\" hidden=\"0\" width=\"8.43\"/>",
    "<col min=\"11\" max=\"12\" collapsed=\"0\" hidden=\"0\" outlineLevel=\"4\" width=\"8.43\"/>",
    "<col min=\"13\" max=\"13\" collapsed=\"0\" width=\"8.43\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
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

test_that("with outlinePr", {

  # create matrix
  t1 <- AirPassengers
  t2 <- do.call(cbind, split(t1, cycle(t1)))
  dimnames(t2) <- dimnames(.preformat.ts(t1))

  wb <- wb_workbook()
  wb$add_worksheet("AirPass")
  wb$add_data("AirPass", t2, rowNames = TRUE)

  wb$worksheets[[1]]$sheetPr <-
    xml_node_create(
      "sheetPr",
      xml_children = c(
        xml_node_create(
          "outlinePr",
          xml_attributes = c(
            summaryBelow = "0",
            summaryRight = "0"
          )
        )
      ))

  # groups will always end on/show the last row. in the example 1950, 1955, and 1960
  wb <- wb_group_rows(wb, "AirPass", 2:3, collapsed = TRUE) # group years < 1950
  wb <- wb_group_rows(wb, "AirPass", 4:8, collapsed = TRUE) # group years 1951-1955
  wb <- wb_group_rows(wb, "AirPass", 9:13)                  # group years 1956-1960

  wb$createCols("AirPass", 13)

  wb <- wb_group_cols(wb, "AirPass", 2:4, collapsed = TRUE)
  wb <- wb_group_cols(wb, "AirPass", 5:7, collapsed = TRUE)
  wb <- wb_group_cols(wb, "AirPass", 8:10, collapsed = TRUE)
  wb <- wb_group_cols(wb, "AirPass", 11:13)


  exp <- c(
    "<col min=\"2\" max=\"2\" collapsed=\"1\" width=\"8.43\"/>",
    "<col min=\"3\" max=\"4\" collapsed=\"1\" hidden=\"1\" outlineLevel=\"1\" width=\"8.43\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr[2:3]
  expect_equal(exp, got)

  exp <- c("", "1")
  got <- wb$worksheets[[1]]$sheet_data$row_attr$outlineLevel[2:3]
  expect_equal(exp, got)

  ### with group levels
  grp_rows <- list(
    "1" = seq(2, 3),
    "2" = seq(4, 8),
    "3" = seq(9, 13)
  )

  grp_cols <- list(
    "1" = seq(2, 4),
    "2" = seq(5, 7),
    "3" = seq(8, 10),
    "4" = seq(11, 13)
  )

  wb <- wb_workbook()
  wb$add_worksheet("AirPass")
  wb$add_data("AirPass", t2, rowNames = TRUE)

  wb$worksheets[[1]]$sheetPr <-
    xml_node_create(
      "sheetPr",
      xml_children = c(
        xml_node_create(
          "outlinePr",
          xml_attributes = c(
            summaryBelow = "0",
            summaryRight = "0"
          )
        )
      ))

  wb$createCols("AirPass", 13)

  wb$group_cols("AirPass", cols = grp_cols)
  wb$group_rows("AirPass", rows = grp_rows)

  exp <- c(
    "<col min=\"2\" max=\"2\" collapsed=\"0\" width=\"8.43\"/>",
    "<col min=\"3\" max=\"4\" collapsed=\"0\" hidden=\"0\" outlineLevel=\"1\" width=\"8.43\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr[2:3]
  expect_equal(exp, got)

  exp <- c("", "1")
  got <- wb$worksheets[[1]]$sheet_data$row_attr$outlineLevel[2:3]
  expect_equal(exp, got)


})
