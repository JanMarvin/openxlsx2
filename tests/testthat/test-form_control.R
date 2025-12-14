test_that("add_form_control works", {

  temp <- temp_xlsx()
  on.exit(unlink(temp), add = TRUE)

  wb <- wb_workbook()$
    # Checkbox
    add_worksheet()$
    add_form_control(dims = "B2")$
    add_form_control(dims = "B3", text = "A text")$
    add_data(dims = "A4", x = 0, col_names = FALSE)$
    add_form_control(dims = "B4", link = "A4")$
    add_data(dims = "A5", x = TRUE, col_names = FALSE)$
    add_form_control(dims = "B5", range = "'Sheet 1'!A5", link = "B5")$
    # Radio
    add_worksheet()$
    add_form_control(dims = "B2", type = "Radio")$
    add_form_control(dims = "B3", type = "Radio", text = "A text")$
    add_data(dims = "A4", x = 0, col_names = FALSE)$
    add_form_control(dims = "B4", type = "Radio", link = "A4")$
    add_data(dims = "A5", x = 1, col_names = FALSE)$
    add_form_control(dims = "B5", type = "Radio")$
    # Drop
    add_worksheet()$
    add_form_control(dims = "B2", type = "Drop")$
    add_form_control(dims = "B3", type = "Drop", text = "A text")$
    add_data(dims = "A4", x = 0, col_names = FALSE)$
    add_form_control(dims = "B4", type = "Drop", link = "A1", range = "D4:D15")$
    add_data(dims = "A5", x = 1, col_names = FALSE)$
    add_form_control(dims = "B5", type = "Drop", link = "'Sheet 3'!D1:D26", range = "A1")$
    add_data(dims = "D1", x = letters)$
    # wide
    add_worksheet()$
    add_form_control(dims = "B2:D3", type = "Checkbox")$
    add_form_control(dims = "B5:D6", type = "Radio")$
    add_form_control(dims = "B8:D9", type = "Drop")

  expect_equal(4L, length(wb$vml))
  expect_equal(4L, length(wb$drawings))
  expect_equal(15L, length(wb$ctrlProps))

  wb$save(temp)

  expect_silent(wb <- wb_load(temp))
  expect_equal(4L, length(wb$vml))
  expect_equal(4L, length(wb$drawings))
  expect_equal(15L, length(wb$ctrlProps))

})
