test_that("Workbook class", {
  expect_null(assert_workbook(wb_workbook()))
})


test_that("wb_set_col_widths", {
# TODO use wb$wb_set_col_widths()

  wb <- wbWorkbook$new()
  wb$add_worksheet("test")
  wb$add_data("test", mtcars)

  # set column width to 12
  expect_silent(wb$set_col_widths("test", widths = 12L, cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"12\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # wrong sheet
  expect_error(wb$set_col_widths("test2", widths = 12L, cols = seq_along(mtcars)))

  # reset the column with, we do not provide an option ot remove the column entry
  expect_silent(wb$set_col_widths("test", cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"8.43\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # create column width for column 25
  expect_silent(wb$set_col_widths("test", cols = "Y", widths = 22))
  expect_equal(
    c("<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"8.43\"/>",
      "<col min=\"12\" max=\"24\" width=\"8.43\"/>",
      "<col min=\"25\" max=\"25\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"22\"/>"),
    wb$worksheets[[1]]$cols_attr
  )

  # a few more errors
  expect_error(wb$set_col_widths("test", cols = "Y", width = 1:2))
  expect_error(wb$set_col_widths("test", cols = "Y", hidden = 1:2))
})


# order -------------------------------------------------------------------

test_that("$set_order() works", {
  wb <- wb_workbook()
  wb$add_worksheet("a")
  wb$add_worksheet("b")
  wb$add_worksheet("c")

  expect_identical(wb$sheetOrder, 1:3)
  exp <- letters[1:3]
  names(exp) <- exp
  expect_identical(wb$get_sheet_names(), exp)

  wb$set_order(3:1)
  expect_identical(wb$sheetOrder, 3:1)
  exp <- letters[3:1]
  names(exp) <- exp
  expect_identical(wb$get_sheet_names(), exp)
})


# sheet names -------------------------------------------------------------

test_that("$set_sheet_names() and $get_sheet_names() work", {
  wb <- wb_workbook()$add_worksheet()$add_worksheet()
  wb$set_sheet_names(new = c("a", "b & c"))

  # return a names character vector
  res <- wb$get_sheet_names()
  exp <- c(a = "a", "b & c" = replace_legal_chars("b & c"))
  expect_identical(res, exp)

  # should be able to check the original values, too
  res <- wb$.__enclos_env__$private$get_sheet_index("b & c")
  expect_identical(res, 2L)
})
