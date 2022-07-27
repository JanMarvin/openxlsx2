test_that("strings_xml", {

  # some sst string
  si <- "<sst><si><t>foo</t></si><si><r><rPr><b/></rPr><t>bar</t></r></si></sst>"

  expect_equal(
    c("foo", "bar"),
    si_to_txt(read_xml(si))
  )

  txt <- "foo"
  expect_equal(
    "<si><t>foo</t></si>",
    txt_to_si(txt, raw = TRUE, skip_control = FALSE)
  )
  expect_equal(
    "<si>\n <t>foo</t>\n</si>\n",
    txt_to_si(txt, raw = FALSE, skip_control = FALSE)
  )

  txt <- "foo "
  expect_equal(
    "<si><t xml:space=\"preserve\">foo </t></si>",
    txt_to_si(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE)
  )
  expect_equal(
    "<si>\n <t xml:space=\"preserve\">foo </t>\n</si>\n",
    txt_to_si(txt, raw = FALSE, no_escapes = FALSE, skip_control = FALSE)
  )

  txt <- "foo&bar"
  expect_equal(
    "<si><t>foo&amp;bar</t></si>",
    txt_to_si(txt, no_escapes = FALSE, skip_control = FALSE)
  )
  expect_equal(
    "<si><t>foo&bar</t></si>",
    txt_to_si(txt, no_escapes = TRUE, skip_control = FALSE)
  )

  txt <- "foo'abcd\037'"
  expect_equal(
    "<si><t>foo'abcd'</t></si>",
    txt_to_si(txt, raw = TRUE, no_escapes = FALSE, skip_control = TRUE)
  )

  is <- c("<is><t>foo</t></is>", "<is><r><rPr><b/></rPr><t>bar</t></r></is>")
  expect_equal(
    c("foo", "bar"),
    is_to_txt(is)
  )

  is <- "<is><t>foo</t>"
  expect_error(is_to_txt(is))


  txt <- "foo"
  expect_equal(
    "<is><t>foo</t></is>",
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE)
  )
  expect_equal(
    "<is>\n <t>foo</t>\n</is>\n",
    txt_to_is(txt, raw = FALSE, no_escapes = FALSE, skip_control = FALSE)
  )

  txt <- "foo "
  expect_equal(
    "<is><t xml:space=\"preserve\">foo </t></is>",
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE)
  )
  expect_equal(
    "<is>\n <t xml:space=\"preserve\">foo </t>\n</is>\n",
    txt_to_is(txt, raw = FALSE, no_escapes = FALSE, skip_control = FALSE)
  )

  txt <- "foo&bar"
  expect_equal(
    "<is><t>foo&amp;bar</t></is>",
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE)
  )
  expect_equal(
    "<is><t>foo&bar</t></is>",
    txt_to_is(txt, raw = TRUE, no_escapes = TRUE, skip_control = FALSE)
  )

  txt <- "foo'abcd\037'"
  expect_equal(
    "<is><t>foo'abcd'</t></is>",
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = TRUE)
  )
})
