test_that("int2col", {

  exp <- LETTERS[1:10]
  got <- int2col(1:10)
  expect_equal(exp, got)

  expect_error(int2col("a"),
               "x must be numeric.")

})

test_that("col2int", {
  expect_null(col2int(NULL))

  expect_equal(1, col2int("a"))
  expect_equal(1, col2int(1))
  expect_error(col2int(list()), "x must be character")

  expect_equal(1, col2int("A"))
  expect_equal(c(1, 3, 4), col2int(c("A", "C:D")))
  expect_equal(c(1, 3, 4, 11), col2int(c("A", "C:D", "K")))
  expect_equal(c(1, 3, 4, 11, 27, 28, 29, 30), col2int(c("A", "C:D", "K", "AA:AD")))

})

test_that("get_cell_refs", {

  exp <- c("B1", "C2", "D3")
  got <- get_cell_refs(data.frame(1:3, 2:4))
  expect_equal(exp, got)

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
