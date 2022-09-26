test_that("get_base_font works", {
  wb <- wb_workbook()
  expect_equal(
    wb$get_base_font(),
    list(
      size = list(val = "11"),
      # should this be "#000000"?
      colour = list(rgb = "FF000000"),
      name = list(val = "Calibri")
    )
  )

  wb$set_base_font( fontSize = 9, fontName = "Arial", fontColour = "red")
  expect_equal(
    wb$get_base_font(),
    list(
      size = list(val = "9"),
      colour = list(rgb = "FFFF0000"),
      name = list(val = "Arial")
    )
  )
})
