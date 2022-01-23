
test_that("Class Style works", {
  expect_error(new_style(), NA)
  expect_error(createStyle(), NA)
  expect_null(assert_style(new_style()))
})

test_that("validate_text_style() works", {
  expect_null(validate_text_style("none"))
  expect_null(validate_text_style(c("none", "bold")))
  expect_identical(validate_text_style(c("underline2", "underline")), "underline2")
})
