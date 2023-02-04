test_that("add_form_control works", {

  temp <- temp_xlsx()

  wb <- wb_workbook()$add_worksheet()$
    # Checkbox
    add_form_control(dims = "B2")$
    add_form_control(dims = "B3", text = "A text")$
    add_data(dims = "A4", x = 0, colNames = FALSE)$
    add_form_control(dims = "B4", link = "A4")$
    add_data(dims = "A5", x = TRUE, colNames = FALSE)$
    add_form_control(dims = "B5", range = "'Sheet 1'!A5", link = "B5")$
    # Radio
    add_worksheet()$
    add_form_control(dims = "B2", type = "Radio")$
    add_form_control(dims = "B3", type = "Radio", text = "A text")$
    add_data(dims = "A4", x = 0, colNames = FALSE)$
    add_form_control(dims = "B4", type = "Radio", link = "A4")$
    add_data(dims = "A5", x = 1, colNames = FALSE)$
    add_form_control(dims = "B5", type = "Radio")$
    # Drop
    add_worksheet()$
    add_form_control(dims = "B2", type = "Drop")$
    add_form_control(dims = "B3", type = "Drop", text = "A text")$
    add_data(dims = "A4", x = 0, colNames = FALSE)$
    add_form_control(dims = "B4", type = "Drop", link = "A1", range = "D4:D15")$
    add_data(dims = "A5", x = 1, colNames = FALSE)$
    add_form_control(dims = "B5", type = "Drop", link = "'Sheet 3'!D1:D26", range = "A1")$
    add_data(dims = "D1", x = letters)

  expect_equal(3L, length(wb$vml))
  expect_equal(3L, length(wb$drawings))
  expect_equal(12L, length(wb$ctrlProps))

  wb$save(temp)

  expect_silent(wb <- wb_load(temp))

})
