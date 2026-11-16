test_that("stri_join handles NA and zero length input", {
  expect_equal(stri_join("a", NA_character_), NA_character_)
  expect_equal(stri_join(c("a", NA), collapse = ""), NA_character_)
  expect_equal(stri_join("a", character()), character())
  expect_equal(stri_join(c("a", "b"), 1:2, sep = "-"), c("a-1", "b-2"))
})

test_that("regex helpers", {
  expect_equal(stri_detect_regex(c("A1", NA, "1A"), "^[A-Z]+[0-9]+$"), c(TRUE, NA, FALSE))

  expect_equal(stri_replace_all_regex("ab", "(a)(b)", "<$2$1>"), "<ba>")
  expect_equal(stri_replace_all_regex("ab", "a", "\\$1"), "$1b")
  expect_error(stri_replace_all_regex("ab", "a", "$0"), "not supported")
  expect_equal(stri_replace_all_regex(character(), "a", "b"), character())
  expect_equal(stri_replace_all_regex(c("a", "b"), c("a", "b"), c("1", "2")), c("1", "2"))
  expect_equal(stri_replace_all_regex(c("ab", "ba"), c("a", "b"), c("1", "2"), vectorize_all = FALSE), c("12", "21"))

  expect_equal(
    stri_locate_all_regex("xyz", "[0-9]")[[1]],
    matrix(NA_integer_, 1L, 2L, dimnames = list(NULL, c("start", "end")))
  )
})

test_that("stri_replace_all_fixed", {
  expect_equal(stri_replace_all_fixed(c("a", "b"), c("a", "b"), c("1", "2")), c("1", "2"))
  expect_equal(stri_replace_all_fixed("ab", c("a", "b"), c("1", "2"), vectorize_all = FALSE), "12")
  expect_equal(stri_replace_all_fixed(character(), "a", "b"), character())

  # native encoded input without a match is returned as UTF-8
  x <- iconv("äx", "UTF-8", "latin1")
  res <- stri_replace_all_fixed(x, "y", "z")
  expect_equal(res, "äx")
  expect_equal(Encoding(res), "UTF-8")
})

test_that("stri_sub and stri_sub_replace", {
  expect_equal(stri_sub("abcdef", -3, -1), "def")
  expect_equal(stri_sub(c("abcdef", "xy"), c(2, -1), c(3, -1)), c("bc", "y"))
  expect_equal(stri_sub("abcdef", NA, 2), NA_character_)

  expect_equal(stri_sub_replace("abcdef", -2, -1, value = "X"), "abcdX")
  expect_equal(stri_sub_replace(c("abc", "xyz"), c(1, -1), c(1, -1), value = c("1", "2")), c("1bc", "xy2"))
  expect_equal(stri_sub_replace("abc", 10, 12, value = "X"), "abcX")
  expect_equal(stri_sub_replace(NA_character_, 1, 1, value = "X"), NA_character_)
  expect_equal(stri_sub_replace("abc", NA, 1, value = "X"), NA_character_)

  # characters above ASCII are counted as characters
  expect_equal(stri_sub_replace("äbc", 2, 1, value = "X"), "äXbc")
  expect_equal(stri_sub_replace("äbc", 2, 2, value = "X"), "äXc")

  expect_equal(stri_sub_replace_all("abc", integer(), integer(), character()), "abc")
})

test_that("stri_locate_last_fixed and stri_detect_fixed", {
  expect_equal(unname(stri_locate_last_fixed("a</x>b</x>c", "</x>")[1, ]), c(7L, 10L))
  expect_true(is.na(stri_locate_last_fixed("ab", "abc")[1, "start"]))
  expect_true(is.na(stri_locate_last_fixed("abc", "x")[1, "start"]))
  expect_true(is.na(stri_locate_last_fixed(NA_character_, "x")[1, "start"]))

  expect_equal(stri_detect_fixed(c("abc", NA, "x"), "b"), c(TRUE, NA, FALSE))
  expect_equal(stri_detect_fixed(character(), "b"), logical())
})

test_that("stri_unique and stri_pad_left", {
  x <- c(iconv("ä", "UTF-8", "latin1"), "ä", "a", NA, "a", NA)
  res <- stri_unique(x)
  expect_equal(res, c("ä", "a", NA))

  expect_equal(stri_pad_left(c("7", NA), 3, "0"), c("007", NA))
  expect_equal(stri_pad_left("7", NA_integer_, "0"), NA_character_)
  expect_equal(stri_pad_left(character(), 3, "0"), character())
})

test_that("stri_order", {
  expect_equal(stri_order(c("b", "a", "C")), c(3L, 2L, 1L))
  opts <- stri_opts_collator(numeric = TRUE)
  x <- c("a10", "a2", "a", "a2", "a02", NA, "a1")
  expect_equal(x[stri_order(x, opts_collator = opts)][1:3], c("a", "a1", "a2"))
  expect_true(is.na(x[stri_order(x, opts_collator = opts)][7]))
  expect_equal(stri_order(c("a", "a1"), opts_collator = opts), 1:2)
  expect_equal(stri_order(c("a1", "a"), opts_collator = opts), 2:1)
  expect_equal(stri_order(c("a1", "a1"), opts_collator = opts), 1:2)
})

test_that("stri_rand_strings", {
  expect_equal(stri_rand_strings(0, 3), character())
  expect_equal(nchar(stri_rand_strings(3, 4)), rep(4L, 3))
})

test_that("stri_read_lines", {
  f <- tempfile()
  on.exit(unlink(f))

  file.create(f)
  expect_equal(stri_read_lines(f), "")

  writeBin(c(as.raw(c(0xEF, 0xBB, 0xBF)), charToRaw("a\nb\n")), f)
  expect_equal(stri_read_lines(f), c("a", "b"))
})

test_that("stri_escape_unicode and stri_unescape_unicode", {
  expect_equal(stri_escape_unicode("aä€\n"), "a\\u00e4\\u20ac\\n")
  expect_equal(stri_escape_unicode("\U0001F600"), "\\U0001f600")
  expect_equal(stri_escape_unicode(c("\"", NA)), c("\\\"", NA))

  expect_equal(stri_unescape_unicode("\\u20ac"), "€")
  expect_equal(stri_unescape_unicode("\\U0001F600"), "\U0001F600")
  expect_equal(stri_unescape_unicode("\\ud83d\\ude00"), "\U0001F600")
  expect_equal(stri_unescape_unicode("\\xe4"), "ä")
  expect_equal(stri_unescape_unicode("\\x{20ac}"), "€")
  expect_equal(stri_unescape_unicode("\\101"), "A")
  expect_equal(stri_unescape_unicode("\\cA"), "\001")
  expect_equal(stri_unescape_unicode("a\\tb\\n"), "a\tb\n")
  expect_equal(stri_unescape_unicode("\\q"), "q")
  expect_equal(stri_unescape_unicode(NA_character_), NA_character_)
  expect_equal(stri_unescape_unicode("abc"), "abc")

  expect_equal(stri_unescape_unicode("\\u"), NA_character_)
  expect_equal(stri_unescape_unicode("\\x"), NA_character_)
  expect_equal(stri_unescape_unicode("\\u00e"), NA_character_)
  expect_equal(stri_unescape_unicode("\\U00110000"), NA_character_)
  expect_equal(stri_unescape_unicode("a\\"), NA_character_)
  expect_error(stri_unescape_unicode("\\u0000"))
  expect_error(stri_unescape_unicode("\\ud800"))
})
