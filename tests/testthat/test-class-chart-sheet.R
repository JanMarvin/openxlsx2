testsetup()

test_that("class wbChartSheet works", {

  wb <- wb_workbook()$add_chartsheet()
  expect_error(wb$add_data(x = mtcars), "Cannot write to chart sheet.")

})

test_that("invalid add_chartsheet arguments leave the workbook unchanged", {

  wb <- wb_workbook()$add_worksheet("A")$add_worksheet("B")

  expect_error(wb$add_chartsheet(tab_color = "invalid"), "Invalid tab_color")
  expect_error(wb$add_chartsheet(zoom = "100"), "zoom must be numeric")
  expect_error(wb$add_chartsheet(zoom = NA_real_), "zoom must be numeric")

  expect_equal(length(wb$workbook$sheets), 2L)

  tmp <- temp_xlsx()
  on.exit(unlink(tmp), add = TRUE)
  expect_silent(wb$save(tmp))
})
