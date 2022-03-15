
test_that("numfmt_class() works", {
  expect_identical(numfmt_class(factor()),                            "factor")
  expect_identical(numfmt_class(Sys.Date()),                          "date")
  expect_identical(numfmt_class(.POSIXct(1)),                         "posix")
  expect_identical(numfmt_class(logical()),                           "logical")
  expect_identical(numfmt_class(character()),                         "character")
  expect_identical(numfmt_class(integer()),                           "integer")
  expect_identical(numfmt_class(numeric()),                           "numeric")
  expect_identical(numfmt_class(structure(1, class = "currency")),    "currency")
  expect_identical(numfmt_class(structure(1, class = "accounting")),  "accounting")
  expect_identical(numfmt_class(structure(1, class = "percentage")),  "percentage")
  expect_identical(numfmt_class(structure(1, class = "scientific")),  "scientific")
  expect_identical(numfmt_class(structure(1, class = "comma")),       "comma")
  expect_identical(numfmt_class(1 ~ 1),                               "formula")
  expect_identical(numfmt_class(structure(1, class = "hyperlink")),   "hyperlink")
  expect_identical(numfmt_class(wb_hyperlink()),                      "hyperlink")
})

test_that("numfmt_class() doesn't work -- when we don't want it", {
  # nested data.frame
  x <- data.frame(a = 1)
  expect_error(numfmt_class(x), NA)
  x[[2]] <- data.frame(b = 1, c = 1)
  expect_error(numfmt_class(x), "data.frame")
})
