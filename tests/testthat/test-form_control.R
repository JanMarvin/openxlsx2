test_that("add_form_control works", {

  wb <- wb_workbook()$add_worksheet()$
    add_form_control(dims = "B2")$
    add_form_control(dims = "B3", text = "A text")$
    add_data(dims = "A4", x = 0, colNames = FALSE)$
    add_form_control(dims = "B4", cell_ref = "A4")$
    add_data(dims = "A5", x = 1, colNames = TRUE)$
    add_form_control(dims = "B5", cell_ref = "'Sheet 1'!A5")

  got <- length(wb$ctrlProps)
  exp <- 4
  expect_equal(exp, got)

})
