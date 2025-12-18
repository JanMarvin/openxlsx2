test_that("int2col", {
  expect_equal(int2col(1:10), LETTERS[1:10])
  expect_error(int2col("a"), "x must be finite and numeric.")
  expect_error(int2col(Inf))
})

test_that("col2int", {
  expect_null(col2int(NULL))

  expect_equal(col2int("a"), 1)
  expect_equal(col2int(1), 1)
  expect_error(col2int(list()), "x must be character")

  expect_equal(col2int("A"), 1)
  expect_equal(col2int(c("A", "C:D")), c(1, 3, 4))
  expect_equal(col2int(c("A", "C:D", "K")), c(1, 3, 4, 11))
  expect_equal(col2int(c("A", "C:D", "K", "AA:AD")), c(1, 3, 4, 11, 27, 28, 29, 30))
  expect_error(col2int(c("a", NA_character_, "c")), "x contains NA")

})

test_that("get_cell_refs", {

  got <- get_cell_refs(data.frame(1:3, 2:4))
  expect_equal(got, c("B1", "C2", "D3"))

  expect_error(get_cell_refs(data.frame("a", "a")),
               "cellCoords must only contain integers")

})

test_that("get_cell_refs() works for a single cell.", {
  expect_equal(get_cell_refs(data.frame(1, 2)), "B1")

  expect_error(get_cell_refs(c(1, 2)))

  expect_error(get_cell_refs(c(1, "a")))
})

test_that("get_cell_refs() works for multiple cells.", {
  expect_equal(get_cell_refs(data.frame(1:3, 2:4)), c("B1", "C2", "D3"))

  expect_error(get_cell_refs(c(1:2, c("a", "b"))))
})

test_that("", {

  op <- options(
    "openxlsx2.maxWidth" = 10,
    "openxlsx2.minWidth" = 8
  )
  on.exit(options(op), add = TRUE)

  got <- calc_col_width(wb_workbook()$get_base_font(), 11)
  expect_equal(got, 10)

  got <- calc_col_width(wb_workbook()$get_base_font(), 7)
  expect_equal(got, 8)

})

test_that("unknown font works", {
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")
  wb$set_base_font(font_name = "Roboto")
  expect_silent(wb$set_col_widths(cols = c(1, 4, 6, 7, 9), widths = c(16, 15, 12, 18, 33)))
})

test_that("row2int works", {
  expect_error(row2int(c("1", NA_character_, "3")), "missings not allowed in rows")
  expect_error(row2int("A"), "missings not allowed in rows")
  expect_equal(row2int(NULL), NULL)
  expect_equal(row2int(character()), integer())
  expect_error(row2int("-1"), "Row exceeds valid range")
  expect_error(row2int("1500000"), "Row exceeds valid range")
  expect_equal(row2int(as.character(1:26)), 1:26)
})
