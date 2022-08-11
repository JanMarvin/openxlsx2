
test_that("class color works", {
  expect_null(assert_color(wb_color()))
})

test_that("get()", {

  exp <- c(rgb = "FF000000")
  expect_false(is_wbColor(exp))

  class(exp) <- c("character", "wbColor")
  got <- wb_color("black")

  expect_true(is_wbColor(got))
  expect_equal(exp, got)

})
