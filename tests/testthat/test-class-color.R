
test_that("class color works", {
  expect_null(assert_color(wb_colour()))
})

test_that("get()", {

  exp <- c(rgb = "FF000000")
  expect_false(is_wbColour(exp))

  class(exp) <- c("character", "wbColour")
  got <- wb_colour("black")

  expect_true(is_wbColour(got))
  expect_equal(exp, got)

})
