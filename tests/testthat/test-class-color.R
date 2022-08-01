
test_that("class color works", {
  expect_null(assert_color(wb_color()))
})

test_that("get()", {

  z <- wb_color("black")

  exp <- c(rgb = "FF000000")
  got <- z$get()
  expect_equal(exp, got)

  exp <- "<color rgb=\"FF000000\"/>"
  got <- z$to_xml()
  expect_equal(exp, got)

})
