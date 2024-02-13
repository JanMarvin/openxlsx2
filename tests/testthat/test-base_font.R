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

  # custom panose values are possible
  wb <- wb_workbook()$
    set_base_font(font_name = "Monaco", font_panose = "xxxxxxxxxxxxxx")
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")

  exp <- "<a:latin typeface=\"Monaco\" panose=\"xxxxxxxxxxxxxx\"/>"
  got <- xml_node(fS, "a:fontScheme", "a:majorFont", "a:latin")
  expect_equal(exp, got)
  got <- xml_node(fS, "a:fontScheme", "a:minorFont", "a:latin")
  expect_equal(exp, got)

  # different font types are possible for panose, not sure how useful this is
  wb <- wb_workbook()$
    set_base_font(font_name = "Arial", font_type = "Italic")
  fS <- xml_node(wb$theme, "a:theme", "a:themeElements", "a:fontScheme")

  exp <- "<a:latin typeface=\"Arial\" panose=\"020B0604020202090204\"/>"
  got <- xml_node(fS, "a:fontScheme", "a:majorFont", "a:latin")
  expect_equal(exp, got)
  got <- xml_node(fS, "a:fontScheme", "a:minorFont", "a:latin")
  expect_equal(exp, got)

})

test_that("hyperlink font size works", {

  wb <- wb_workbook()$
    set_base_font(font_size = 13, font_name = "Monaco")$
    add_worksheet()$
    add_formula(x = create_hyperlink(text = "foo", file = "bar"))

  exp <- c(
    "<font><color theme=\"1\"/><family val=\"2\"/><name val=\"Monaco\"/><scheme val=\"minor\"/><sz val=\"13\"/></font>",
    "<font><color theme=\"10\"/><name val=\"Monaco\"/><sz val=\"13\"/><u val=\"single\"/></font>"
  )
  got <- wb$styles_mgr$styles$fonts
  expect_equal(exp, got)

})
