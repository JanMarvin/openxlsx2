test_that("clean_worksheet_name", {

  exp <- "test that"
  got <- clean_worksheet_name("test\that")
  expect_equal(exp, got)

  exp <- c("a b", "a b", "a b", "a b", "a b", "a b", "a b")
  got <- clean_worksheet_name(c("a\b", "a/b", "a?b", "a*b", "a:b", "a[b", "a]b"))
  expect_equal(exp, got)

})
