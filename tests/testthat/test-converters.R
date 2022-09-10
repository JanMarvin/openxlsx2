test_that("int2col", {

  exp <- LETTERS[1:10]
  got <- int2col(1:10)
  expect_equal(exp, got)

  expect_error(int2col("a"),
               "x must be numeric.")

})

test_that("col2int", {

  expect_equal(1, col2int("a"))
  expect_equal(1, col2int(1))
  expect_error(col2int(list()), "x must be character")

})

test_that("get_cell_refs", {

  exp <- c("B1", "C2", "D3")
  got <- get_cell_refs(data.frame(1:3, 2:4))
  expect_equal(exp, got)

  expect_error(get_cell_refs(data.frame("a", "a")),
               "cellCoords must only contain integers")

})
