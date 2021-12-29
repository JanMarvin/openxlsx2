
test_that("assert_class() works", {
  expect_null(assert_class(2L, "integer"))
  expect_null(assert_class(2, "numeric"))
  expect_null(assert_class(2, c("numeric", "integer")))
  expect_null(assert_class("2", "character"))
  expect_null(assert_class(Sys.Date(), "Date"))
  expect_null(assert_class(Sys.time(), "POSIXt"))
})

test_that("class utils", {
  expect_null(assert_chart_sheet(new_chart_sheet()))
  expect_null(assert_comment(    new_comment()))
  expect_null(assert_hyperlink(  new_hyperlink()))
  expect_null(assert_sheet_data( new_sheet_data()))
  expect_null(assert_style(      new_style()))
  expect_null(assert_workbook(   new_workbook()))
  expect_null(assert_worksheet(  new_worksheet()))
})

test_that("match_oneof() works", {
  expect_identical(match_oneof(1:4, 3:4), 3L)
  expect_identical(match_oneof(1:4, 3:4, several = TRUE), 3:4)
  expect_null(match_oneof(NULL, 1:3, or_null = TRUE))
  expect_error(match_oneof("d", letters[1:3]))
  expect_error(match_oneof("d", letters[1:3], several = TRUE))
})
