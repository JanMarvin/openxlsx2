
test_that("class utils", {
  expect_null(assert_chart_sheet(new_chart_sheet()))
  expect_null(assert_comment(    new_comment()))
  expect_null(assert_hyperlink(  new_hyperlink()))
  expect_null(assert_sheet_data( new_sheet_data()))
  expect_null(assert_style(      new_style()))
  expect_null(assert_workbook(   new_workbook()))
  expect_null(assert_worksheet(  new_worksheet()))
})
