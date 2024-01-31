test_that("int2col", {

  expect_equal(int2col(1:10), LETTERS[1:10])
  expect_error(int2col("a"), "x must be numeric.")
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
