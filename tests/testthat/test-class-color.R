
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

test_that("tabColour can be wb_colour()", {
  expect_silent(
    wb_workbook()$
      # wb_colour
      add_worksheet(tabColour = wb_colour("green"))$
      add_chartsheet(tabColour = wb_colour("green"))$
      # colour name
      add_worksheet(tabColour = "green")$
      add_chartsheet(tabColour = "green")
  )
})
