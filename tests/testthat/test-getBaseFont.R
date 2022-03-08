test_that("getBaseFont works", {
  wb <- wb_workbook()
  expect_equal(
    wb$getBaseFont(),
    list(
      size = list(val = "11"),
      # should this be "#000000"?
      colour = list(rgb = "FF000000"),
      name = list(val = "Calibri")
    )
  )

  modifyBaseFont(wb, fontSize = 9, fontName = "Arial", fontColour = "red")
  expect_equal(
    wb$getBaseFont(),
    list(
      size = list(val = "9"),
      colour = list(rgb = "FFFF0000"),
      name = list(val = "Arial")
    )
  )
})
