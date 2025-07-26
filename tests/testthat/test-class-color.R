testsetup()

test_that("class color works", {
  expect_null(assert_color(wb_colour()))
})

test_that("get()", {

  exp <- c(rgb = "FF000000")
  expect_false(is_wbColour(exp))

  class(exp) <- c("wbColour", "character")
  got <- wb_colour("black")

  expect_true(is_wbColour(got))
  expect_equal(exp, got)

})

test_that("tabColour can be wb_colour()", {
  expect_silent(
    wb_workbook()$
      # wb_colour
      add_worksheet(tab_colour = wb_colour("green"))$
      add_chartsheet(tab_colour = wb_colour("green"))$
      # colour name
      add_worksheet(tab_colour = "green")$
      add_chartsheet(tab_colour = "green")
  )
})

test_that("treat color and colour equally", {

  wb_color <- wb_workbook()$
    add_worksheet(tab_color = "green")$
    add_fill(color = wb_color("blue"))$
    add_border(
      dims   = "G12:H13",
      left_color   = wb_color("red"),
      right_color  = wb_color("blue"),
      top_color    = wb_color("green"),
      bottom_color = wb_color("yellow")
    )

  wb_colour <- wb_workbook()$
    add_worksheet(tab_colour = "green")$
    add_fill(colour = wb_colour("blue"))$
    add_border(
      dims   = "G12:H13",
      left_colour   = wb_colour("red"),
      right_colour  = wb_colour("blue"),
      top_colour    = wb_colour("green"),
      bottom_colour = wb_colour("yellow")
    )

  expect_equal(
    wb_color$styles_mgr$styles,
    wb_colour$styles_mgr$styles
  )

})
