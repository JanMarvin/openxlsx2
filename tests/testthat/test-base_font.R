testsetup()

test_that("get_base_font works", {
  wb <- wb_workbook()
  expect_equal(
    wb$get_base_font(),
    list(
      size = list(val = "11"),
      # should this be "#000000"?
      color = list(theme = "1"),
      name = list(val = "Aptos Narrow")
    )
  )

  wb$set_base_font(fontSize = 9, fontName = "Arial", fontColour = wb_colour("red"))
  expect_equal(
    wb$get_base_font(),
    list(
      size = list(val = "9"),
      color = list(rgb = "FFFF0000"),
      name = list(val = "Arial")
    )
  )
})

test_that("wb_set_base_font() actually alters the base font", {

  wb <- wb_workbook()$add_worksheet()

  # per default a workbook has no theme until it is written
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")
  expect_equal(character(), fS)

  # if we alter the base font, a theme is created
  wb$set_base_font(font_name = "Calibri")
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")

  exp <- "<a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/>"
  got <- xml_node(fS, "a:fontScheme", "a:majorFont", "a:latin")
  expect_equal(exp, got)
  got <- xml_node(fS, "a:fontScheme", "a:minorFont", "a:latin")
  expect_equal(exp, got)


  # if we request a theme, this is present
  wb <- wb_workbook(theme = "Old Office Theme")$add_worksheet()
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")

  exp <- "<a:latin typeface=\"Cambria\"/>"
  got <- xml_node(fS, "a:fontScheme", "a:majorFont", "a:latin")
  expect_equal(exp, got)

  exp <- "<a:latin typeface=\"Calibri\"/>"
  got <- xml_node(fS, "a:fontScheme", "a:minorFont", "a:latin")
  expect_equal(exp, got)

  # this can also be altered
  wb$set_base_font(font_name = "Aptos Narrow")
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")

  exp <- "<a:latin typeface=\"Aptos Narrow\" panose=\"020B0004020202020204\"/>"
  got <- xml_node(fS, "a:fontScheme", "a:majorFont", "a:latin")
  expect_equal(exp, got)
  got <- xml_node(fS, "a:fontScheme", "a:minorFont", "a:latin")
  expect_equal(exp, got)

  wb$set_base_font(font_name = "Calibri")
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")

  # only the font_size is altered
  wb <- wb_workbook()$add_worksheet()$set_base_font(font_size = 16)
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")
  expect_equal(character(), fS)

})
