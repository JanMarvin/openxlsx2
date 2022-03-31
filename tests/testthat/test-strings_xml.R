test_that("strings_xml", {

  # some sst string
  si <- "<sst><si><t>foo</t></si><si><r><rPr><b/></rPr><t>bar</t></r></si></sst>"

  expect_equal(
    c("foo", "bar"),
    openxlsx2:::si_to_txt(read_xml(si))
  )

  txt <- "foo"
  expect_equal(
    "<si><t>foo</t></si>",
    openxlsx2:::txt_to_si(txt, raw = TRUE)
  )
  expect_equal(
    "<si>\n <t>foo</t>\n</si>\n",
    openxlsx2:::txt_to_si(txt, raw = FALSE)
  )

  txt <- "foo&bar"
  expect_equal(
    "<si><t>foo&amp;bar</t></si>",
    openxlsx2:::txt_to_si(txt, no_escapes = FALSE)
  )
  expect_equal(
    "<si><t>foo&bar</t></si>",
    openxlsx2:::txt_to_si(txt, no_escapes = TRUE)
  )

  is <- c("<is><t>foo</t></is>", "<is><r><rPr><b/></rPr><t>bar</t></r></is>")
  expect_equal(
    c("foo", "bar"),
    openxlsx2:::is_to_txt(is)
  )

  is <- "<is><t>foo</t>"
  expect_error(openxlsx2:::is_to_txt(is))


  txt <- "foo"
  expect_equal(
    "<is><t>foo</t></is>",
    openxlsx2:::txt_to_is(txt, raw = TRUE, no_escapes = FALSE)
  )
  expect_equal(
    "<is>\n <t>foo</t>\n</is>\n",
    openxlsx2:::txt_to_is(txt, raw = FALSE, no_escapes = FALSE)
  )

  txt <- "foo&bar"
  expect_equal(
    "<is><t>foo&amp;bar</t></is>",
    openxlsx2:::txt_to_is(txt, raw = TRUE, no_escapes = FALSE)
  )
  expect_equal(
    "<is><t>foo&bar</t></is>",
    openxlsx2:::txt_to_is(txt, raw = TRUE, no_escapes = TRUE)
  )
})
