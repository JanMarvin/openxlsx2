test_that("Provide tests for single get_cell_refs", {
  expect_equal(get_cell_refs(data.frame(1, 2)), "B1")

  expect_error(get_cell_refs(c(1, 2)))

  expect_error(get_cell_refs(c(1, "a")))
})

test_that("Provide tests for multiple get_cell_refs", {
  expect_equal(get_cell_refs(data.frame(1:3, 2:4)), c("B1", "C2", "D3"))

  expect_error(get_cell_refs(c(1:2, c("a", "b"))))
})
