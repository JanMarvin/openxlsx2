
test_that("genBaseColStyle() works", {
  x <- genBaseColStyle(1)
  expect_length(x, 2L)
  expect_type(x, "list")
  expect_s4_class(x$style, "Style")

  expect_error(genBaseColStyle("date"), NA)
  expect_error(genBaseColStyle("posixlt"), NA)
  expect_error(genBaseColStyle("currency"), NA)
  expect_error(genBaseColStyle("accounting"), NA)
  expect_error(genBaseColStyle("hyperlink"), NA)
  expect_error(genBaseColStyle("percentage"), NA)
  expect_error(genBaseColStyle("scientific"), NA)
  expect_error(genBaseColStyle("comma"), NA)
  expect_error(genBaseColStyle("numeric"), NA)
  expect_error(genBaseColStyle("0.00"), NA)
  expect_error(genBaseColStyle("something-else"), NA)
})
