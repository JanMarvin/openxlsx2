test_that("clean_worksheet_name", {

  exp <- "test that"
  got <- clean_worksheet_name("test\that")
  expect_equal(exp, got)

  exp <- c("a b", "a b", "a b", "a b", "a b", "a b", "a b")
  got <- clean_worksheet_name(c("a\b", "a/b", "a?b", "a*b", "a:b", "a[b", "a]b"))
  expect_equal(exp, got)

})

test_that("control characters do not end up in the xml files", {

  exp <- c("ab", "ab&#9;c&#10;d&#13;e", "a\tb\nc")
  got <- strip_xml_control_chars(c("a\u0001b", "a&#01;b&#x1f;&#9;c&#10;d&#13;e&#x0B;", "a\tb\nc"))
  expect_equal(got, exp)

  wb <- wb_workbook(title = "a\u0001b")$
    add_worksheet()$
    add_data(x = 1)$
    add_font(dims = "A1", name = "Ari\u0002al")$
    add_data_validation(dims = "B1", type = "whole", operator = "between",
                        value = c(1, 9), prompt = "pro\u0003mpt")

  temp <- temp_xlsx()
  on.exit(unlink(temp), add = TRUE)
  wb$save(temp)

  tmp_dir <- temp_dir("control_chars")
  on.exit(unlink(tmp_dir, recursive = TRUE), add = TRUE)
  utils::unzip(temp, exdir = tmp_dir)

  fls <- list.files(tmp_dir, pattern = "\\.xml$", recursive = TRUE, full.names = TRUE)
  for (fl in fls) {
    raw <- readBin(fl, "raw", file.size(fl))
    expect_false(any(raw < as.raw(0x20) & !raw %in% as.raw(c(0x09, 0x0a, 0x0d))), info = fl)
    txt <- rawToChar(raw)
    expect_false(grepl("&#0*1;|&#0*2;|&#0*3;", txt), info = fl)
  }

  expect_silent(wb_load(temp))

})

test_that("control characters are removed from formulas and the theme", {

  wb <- wb_workbook()$
    set_base_font(font_name = "Ari\u0001al")$
    add_worksheet()$
    add_formula(x = "\"a\u0001b\"", dims = "A1")

  temp <- temp_xlsx()
  on.exit(unlink(temp), add = TRUE)
  wb$save(temp)

  tmp_dir <- temp_dir("control_chars_fml")
  on.exit(unlink(tmp_dir, recursive = TRUE), add = TRUE)
  utils::unzip(temp, exdir = tmp_dir)

  for (fl in file.path(tmp_dir, c("xl/worksheets/sheet1.xml", "xl/theme/theme1.xml"))) {
    raw <- readBin(fl, "raw", file.size(fl))
    expect_false(any(raw < as.raw(0x20) & !raw %in% as.raw(c(0x09, 0x0a, 0x0d))), info = fl)
    expect_false(grepl("&#0*1;", rawToChar(raw)), info = fl)
  }

})
