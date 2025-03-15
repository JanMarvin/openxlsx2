test_that("strings_xml", {

  # some sst string
  si <- "<sst><si><t>foo</t></si><si><r><rPr><b/></rPr><t>bar</t></r></si></sst>"
  expect_equal(
    xml_si_to_txt(read_xml(si)),
    c("foo", "bar")
  )

  txt <- "foo"
  expect_equal(
    txt_to_si(txt, raw = TRUE, skip_control = FALSE),
    "<si><t>foo</t></si>"
  )
  expect_equal(
    txt_to_si(txt, raw = FALSE, skip_control = FALSE),
    "<si>\n <t>foo</t>\n</si>\n"
  )

  txt <- "foo "
  expect_equal(
    txt_to_si(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE),
    "<si><t xml:space=\"preserve\">foo </t></si>"
  )
  expect_equal(
    txt_to_si(txt, raw = FALSE, no_escapes = FALSE, skip_control = FALSE),
    "<si>\n <t xml:space=\"preserve\">foo </t>\n</si>\n"
  )

  txt <- "foo&bar"
  expect_equal(
    txt_to_si(txt, no_escapes = FALSE, skip_control = FALSE),
    "<si><t>foo&amp;bar</t></si>"
  )
  expect_equal(
    txt_to_si(txt, no_escapes = TRUE, skip_control = FALSE),
    "<si><t>foo&bar</t></si>"
  )

  txt <- "foo'abcd\037'"
  expect_equal(
    txt_to_si(txt, raw = TRUE, no_escapes = FALSE, skip_control = TRUE),
    "<si><t>foo'abcd'</t></si>"
  )

  is <- c("<is><t>foo</t></is>", "<is><r><rPr><b/></rPr><t>bar</t></r></is>")
  expect_equal(
    is_to_txt(is),
    c("foo", "bar")
  )

  is <- "<is><t>foo</t>"
  expect_error(is_to_txt(is))

  exp <- c("a", "", "")
  got <- is_to_txt(c("<is><t>a</t></is>", "", "<is><t/></is>"))
  expect_equal(exp, got)

  txt <- "foo"
  expect_equal(
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE),
    "<is><t>foo</t></is>"
  )
  expect_equal(
    txt_to_is(txt, raw = FALSE, no_escapes = FALSE, skip_control = FALSE),
    "<is>\n <t>foo</t>\n</is>\n"
  )

  txt <- "foo "
  expect_equal(
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE),
    "<is><t xml:space=\"preserve\">foo </t></is>"
  )
  expect_equal(
    txt_to_is(txt, raw = FALSE, no_escapes = FALSE, skip_control = FALSE),
    "<is>\n <t xml:space=\"preserve\">foo </t>\n</is>\n"
  )

  txt <- "foo&bar"
  expect_equal(
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = FALSE),
    "<is><t>foo&amp;bar</t></is>"
  )
  expect_equal(
    txt_to_is(txt, raw = TRUE, no_escapes = TRUE, skip_control = FALSE),
    "<is><t>foo&bar</t></is>"
  )

  txt <- "foo'abcd\037'"
  expect_equal(
    txt_to_is(txt, raw = TRUE, no_escapes = FALSE, skip_control = TRUE),
    "<is><t>foo'abcd'</t></is>"
  )

  amp <- temp_xlsx()
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(dims = "A1", x = "A & B")$
    save(amp)

  got <- wb_to_df(amp, col_names = FALSE)[1, 1]
  expect_equal(got, "A & B")

  # a couple of saves and loads later ...
  got <- wb$worksheets[[1]]$sheet_data$cc$is
  expect_equal(got, "<is><t>A &amp; B</t></is>")

  wb <- wb_load(amp)

  got <- wb$worksheets[[1]]$sheet_data$cc$is
  expect_equal(got, "<is><t>A &amp; B</t></is>")

  wb$save(amp)

  wb <- wb_load(amp)
  got <- wb$worksheets[[1]]$sheet_data$cc$is
  expect_equal(got, "<is><t>A &amp; B</t></is>")

  got <- txt_to_is('foo <em>bar</em>')
  exp <- "<is><t>foo &lt;em&gt;bar&lt;/em&gt;</t></is>"
  expect_equal(got, exp)

  # exception to the rule: it is not possible to write characters starting with "<r/>" or "<r>""
  got <- txt_to_is("<r>foo</r>")
  exp <- "<is><r>foo</r></is>"
  expect_equal(got, exp)

  got <- txt_to_is(fmt_txt('<r>foo</r>'))
  exp <- "<is><r><rPr/><t>&lt;r&gt;foo&lt;/r&gt;</t></r></is>"
  expect_equal(got, exp)

  got <- txt_to_is('<red>foo</red>')
  exp <- "<is><t>&lt;red&gt;foo&lt;/red&gt;</t></is>"
  expect_equal(got, exp)

  got <- txt_to_is('foo</r>')
  exp <- "<is><t>foo&lt;/r&gt;</t></is>"
  expect_equal(got, exp)

  got <- txt_to_is(' <r>foo</r>')
  exp <- "<is><t xml:space=\"preserve\"> &lt;r&gt;foo&lt;/r&gt;</t></is>"
  expect_equal(got, exp)

})
