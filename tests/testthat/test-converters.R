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

test_that("as_character", {

  chr <- "433256.529740090366"
  num <-  433256.529740090366

  op <- openxlsx2_options()
  on.exit(options(op), add = TRUE)

  exp <- "433256.5297400903655216"
  got <- as_character(num)

  input <- c(1, NA, 3)

  exp <- c("1", NA_character_, "3")
  got <- as_character(input)
  expect_equal(exp, got)

})
