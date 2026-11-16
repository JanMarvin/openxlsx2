# Replacements for the stringi functions used in openxlsx2. They keep the
# stringi names and arguments, so call sites did not change. Only the argument
# combinations used in this package are covered. The functions are documented
# with @noRd, the notes below list where they differ from stringi.
#
# Encoding: stringi converts every input to UTF-8. The functions here do the
# same with enc2utf8(), so strings marked as latin1 or in the native encoding
# of a non UTF-8 session are converted correctly. Non ASCII strings are then
# processed in UTF-8 mode in any locale and the results are marked as UTF-8.
# stringi itself reads unmarked native strings as UTF-8, which gives garbage
# in a latin1 or cp1252 session.
#
# A byte order mark (U+FEFF) at the start of a string is dropped by stringi in
# several functions (join, replace, sub, ...) and kept here.

#' Join strings
#'
#' Replacement for `stringi::stri_join()`. Same as in stringi, `NA` makes the
#' whole result `NA` (`paste()` would write "NA") and a zero length argument
#' gives a zero length result. Arguments that are not character are converted
#' with `as.character()`.
#' @param ... character vectors
#' @param sep separator between the vectors
#' @param collapse optional separator to join the result into one string
#' @examples
#' stri_join("a", c("x", "y"), sep = "-")
#' #> [1] "a-x" "a-y"
#' stri_join(c("a", NA), "b")
#' #> [1] "ab" NA
#' stri_join(c("a", "b"), collapse = ", ")
#' #> [1] "a, b"
#' stri_join("a", character())
#' #> character(0)
#' @noRd
stri_join <- function(..., sep = "", collapse = NULL) {
  args <- list(...)
  obj <- vapply(args, is.object, NA)
  if (any(obj)) args[obj] <- lapply(args[obj], as.character)
  if (is.null(collapse)) return(stri_join_cpp(args, sep))

  # NA is contagious, like in stringi. paste() would write "NA"
  if (length(args) == 1L && is.character(args[[1L]])) return(stri_collapse_cpp(args[[1L]], collapse))
  stri_collapse_cpp(stri_join_cpp(args, sep), collapse)
}

#' Prepare a pattern for PCRE
#'
#' PCRE only accepts code points above 255 (`\x{100}` and up) in UTF mode. R
#' switches to UTF mode when the pattern contains a non ASCII character, so a
#' harmless empty group with one is appended to patterns that need it.
#' @param x character vector of patterns
#' @examples
#' pcre_pattern("[0-9]+")
#' #> [1] "[0-9]+"
#' endsWith(pcre_pattern("\\x{20ac}"), "){0}")
#' #> [1] TRUE
#' @noRd
pcre_pattern <- function(x) {
  x <- enc2utf8(as.character(x))
  hi <- grepl("\\\\x\\{[0-9a-fA-F]{3,}\\}", x)
  x[hi] <- paste0(x[hi], "(?:\u00e4){0}")
  x
}

#' Detect a fixed string
#'
#' Replacement for `stringi::stri_detect_fixed()`. Only a single, non empty
#' pattern is supported. `NA` in `str` gives `NA`.
#' @param str character vector
#' @param pattern single string
#' @examples
#' stri_detect_fixed(c("a&b", "ab", NA), "&")
#' #> [1]  TRUE FALSE    NA
#' @noRd
stri_detect_fixed <- function(str, pattern) {
  if (is.object(str)) str <- as.character(str)
  if (is.object(pattern)) pattern <- as.character(pattern)
  stri_detect_fixed_cpp(str, pattern)
}

#' Detect a regular expression
#'
#' Replacement for `stringi::stri_detect_regex()`. stringi uses ICU regular
#' expressions, this uses PCRE (`perl = TRUE`). For the patterns used in this
#' package both behave the same. Known differences: in PCRE `\d`, `\s` and
#' `\w` only match ASCII characters and lookbehind must have a fixed length.
#' ICU treats carriage return, U+0085, U+2028 and U+2029 as line terminators
#' next to the newline, so `.` does not match them and `$` can match in front
#' of them. PCRE only knows the newline. The same applies to all other
#' `*_regex` replacements below.
#' @param str character vector
#' @param pattern single regular expression
#' @examples
#' stri_detect_regex(c("A1", "1A", NA), "^[A-Z]+[0-9]+$")
#' #> [1]  TRUE FALSE    NA
#' @noRd
stri_detect_regex <- function(str, pattern) {
  str <- enc2utf8(as.character(str))
  res <- grepl(pcre_pattern(pattern), str, perl = TRUE)
  if (anyNA(str)) res[is.na(str)] <- NA
  res
}

#' Translate a replacement string
#'
#' ICU and PCRE write replacement strings differently. In ICU `$1` is a group
#' and a backslash escapes the next character. This turns such a string into
#' one for `gsub(perl = TRUE)`. `$0` is not supported and raises an error.
#' @param x single replacement string
#' @examples
#' icu_replacement("$1-$2")
#' #> [1] "\\1-\\2"
#' icu_replacement("\\$1")
#' #> [1] "$1"
#' @noRd
icu_replacement <- function(x) {
  x <- enc2utf8(x)
  if (is.na(x) || !grepl("[\\\\$]", x)) return(x)
  if (grepl("(^|[^\\\\])(\\\\\\\\)*\\$0", x)) stop("`$0` is not supported in replacement strings", call. = FALSE)

  tokens <- regmatches(x, gregexpr("(?s)\\\\.|\\$[0-9]|[^\\\\$]+|.", x, perl = TRUE))[[1]]

  res <- vapply(tokens, function(tok) {
    if (grepl("^\\$[0-9]$", tok)) return(sprintf("\\%s", substr(tok, 2L, 2L)))
    if (nchar(tok) == 2L && substr(tok, 1L, 1L) == "\\") tok <- substr(tok, 2L, 2L)
    gsub("\\", "\\\\", tok, fixed = TRUE)
  }, NA_character_, USE.NAMES = FALSE)

  paste(res, collapse = "")
}

#' Replace regular expression matches
#'
#' Replacement for `stringi::stri_replace_all_regex()`. Replacement strings
#' follow the ICU syntax. With `vectorize_all = FALSE` every pattern is applied
#' to every string, one after the other.
#' @param str character vector
#' @param pattern regular expression(s)
#' @param replacement replacement string(s)
#' @param vectorize_all recycle patterns and replacements along `str` (TRUE) or
#'   apply them one after the other (FALSE)
#' @examples
#' stri_replace_all_regex("a1b22", "([a-z])([0-9]+)", "$2$1")
#' #> [1] "1a22b"
#' stri_replace_all_regex("abc", c("a", "b"), c("b", "c"), vectorize_all = FALSE)
#' #> [1] "ccc"
#' @noRd
stri_replace_all_regex <- function(str, pattern, replacement, vectorize_all = TRUE) {
  str <- enc2utf8(as.character(str))
  pattern <- pcre_pattern(pattern)
  replacement <- vapply(as.character(replacement), icu_replacement, NA_character_, USE.NAMES = FALSE)

  if (!length(str) || !length(pattern) || !length(replacement)) return(character())

  if (!vectorize_all) {
    # every pattern is applied to every string, one after the other
    n <- max(length(pattern), length(replacement))
    pattern <- rep_len(pattern, n)
    replacement <- rep_len(replacement, n)
    for (i in seq_len(n)) str <- gsub(pattern[i], replacement[i], str, perl = TRUE)
    return(str)
  }

  if (length(pattern) == 1L && length(replacement) == 1L) {
    return(gsub(pattern, replacement, str, perl = TRUE))
  }

  n <- max(length(str), length(pattern), length(replacement))
  str <- rep_len(str, n)
  pattern <- rep_len(pattern, n)
  replacement <- rep_len(replacement, n)
  res <- character(n)
  for (i in seq_len(n)) res[i] <- gsub(pattern[i], replacement[i], str[i], perl = TRUE)
  res
}

#' Replace fixed strings
#'
#' Replacement for `stringi::stri_replace_all_fixed()`, written in C++. Unchanged
#' elements are returned as they are. Patterns must not be empty.
#' @inheritParams stri_replace_all_regex
#' @examples
#' stri_replace_all_fixed(c("a&b", NA), "&", "&amp;")
#' #> [1] "a&amp;b" NA
#' stri_replace_all_fixed("abc", c("a", "b"), c("b", "c"), vectorize_all = FALSE)
#' #> [1] "ccc"
#' @noRd
stri_replace_all_fixed <- function(str, pattern, replacement, vectorize_all = TRUE) {
  if (is.character(str) && is.character(pattern) && is.character(replacement) &&
      length(str) && length(pattern) == 1L && length(replacement) == 1L) {
    return(stri_replace_all_fixed_cpp(str, pattern, replacement))
  }

  str <- as.character(str)
  pattern <- as.character(pattern)
  replacement <- as.character(replacement)

  if (!length(str) || !length(pattern) || !length(replacement)) return(character())

  # every pattern is applied to every string, one after the other
  if (!vectorize_all || (length(pattern) == 1L && length(replacement) == 1L)) {
    n <- max(length(pattern), length(replacement))
    return(stri_replace_all_fixed_cpp(str, rep_len(pattern, n), rep_len(replacement, n)))
  }

  n <- max(length(str), length(pattern), length(replacement))
  str <- rep_len(str, n)
  pattern <- rep_len(pattern, n)
  replacement <- rep_len(replacement, n)
  res <- character(n)
  for (i in seq_len(n)) res[i] <- stri_replace_all_fixed_cpp(str[i], pattern[i], replacement[i])
  res
}

#' Extract all matches
#'
#' Replacement for `stringi::stri_extract_all_regex()`. `gregexpr()` is slow on
#' long vectors, so only elements that contain a match are passed to it and the
#' matches are cut out of all strings at once. Elements without a match are
#' `NA` or an empty vector with `omit_no_match = TRUE`, same as in stringi.
#' Patterns that can match the empty string are not supported, stringi reports
#' an additional empty match at the end of the string.
#' @param str character vector
#' @param pattern single regular expression
#' @param omit_no_match return `character(0)` instead of `NA` if nothing matched
#' @examples
#' stri_extract_all_regex(c("a1b22", "x"), "[0-9]+")
#' #> [[1]]
#' #> [1] "1"  "22"
#' #>
#' #> [[2]]
#' #> [1] NA
#' #>
#' stri_extract_all_regex(c("a1b22", "x"), "[0-9]+", omit_no_match = TRUE)
#' #> [[1]]
#' #> [1] "1"  "22"
#' #>
#' #> [[2]]
#' #> character(0)
#' #>
#' @noRd
stri_extract_all_regex <- function(str, pattern, omit_no_match = FALSE) {
  str <- enc2utf8(as.character(str))
  pattern <- pcre_pattern(pattern)

  # gregexpr() and regmatches() are slow for many strings. only strings with a
  # match are passed to them, and the matches are extracted at once
  hit <- which(grepl(pattern, str, perl = TRUE))
  res <- rep(list(if (omit_no_match) character() else NA_character_), length(str))

  if (length(hit)) {
    m <- gregexpr(pattern, str[hit], perl = TRUE)
    start <- unlist(m, use.names = FALSE)
    len <- unlist(lapply(m, attr, "match.length"), use.names = FALSE)
    id <- rep.int(seq_along(hit), lengths(m))
    found <- start > 0L
    res[hit] <- unname(split(substring(str[hit][id[found]], start[found], start[found] + len[found] - 1L),
                             factor(id[found], seq_along(hit))))
  }

  res[is.na(str)] <- list(NA_character_)
  res
}

#' Extract the first match
#'
#' Replacement for `stringi::stri_extract_first_regex()`. The match positions
#' come from `regexpr()`, the strings are cut in C++ to avoid copies of long
#' vectors.
#' @param str character vector
#' @param pattern single regular expression
#' @examples
#' stri_extract_first_regex(c("a1b22", "x", NA), "[0-9]+")
#' #> [1] "1" NA  NA
#' @noRd
stri_extract_first_regex <- function(str, pattern) {
  str <- enc2utf8(as.character(str))
  m <- regexpr(pcre_pattern(pattern), str, perl = TRUE)
  stri_extract_first_cpp(str, m, attr(m, "match.length"))
}

#' Match the first occurrence including groups
#'
#' Replacement for `stringi::stri_match_first_regex()`. Returns the same matrix:
#' one row per string, the full match in the first column and one column per
#' group. Groups that did not take part in the match are `NA`.
#' @param str character vector
#' @param pattern single regular expression
#' @examples
#' stri_match_first_regex(c("A1:B2", "C3", "x"), "([A-Z]+)([0-9]+)(:([A-Z]+)([0-9]+))?")
#' #>      [,1]    [,2] [,3] [,4]  [,5] [,6]
#' #> [1,] "A1:B2" "A"  "1"  ":B2" "B"  "2"
#' #> [2,] "C3"    "C"  "3"  NA    NA   NA
#' #> [3,] NA      NA   NA   NA    NA   NA
#' @noRd
stri_match_first_regex <- function(str, pattern) {
  str <- enc2utf8(as.character(str))
  m <- regexpr(pcre_pattern(pattern), str, perl = TRUE)
  start <- attr(m, "capture.start")
  len <- attr(m, "capture.length")
  ngroups <- if (is.null(start)) 0L else ncol(start)

  res <- matrix(NA_character_, nrow = length(str), ncol = ngroups + 1L)
  hit <- !is.na(m) & m > 0L
  res[hit, 1L] <- substring(str[hit], m[hit], m[hit] + attr(m, "match.length")[hit] - 1L)

  for (j in seq_len(ngroups)) {
    # groups that did not take part in the match are NA, empty matches are ""
    set <- hit & start[, j] > 0L
    res[set, j + 1L] <- substring(str[set], start[set, j], start[set, j] + len[set, j] - 1L)
  }

  res
}

#' Locate all matches
#'
#' Replacement for `stringi::stri_locate_all_regex()`. A list with one matrix
#' (columns `start` and `end`, in characters) per string. Without a match the
#' matrix has one row of `NA`, as in stringi.
#' @param str character vector
#' @param pattern single regular expression
#' @examples
#' stri_locate_all_regex("a1b22", "[0-9]+")
#' #> [[1]]
#' #>      start end
#' #> [1,]     2   2
#' #> [2,]     4   5
#' #>
#' stri_locate_all_regex("abc", "[0-9]+")
#' #> [[1]]
#' #>      start end
#' #> [1,]    NA  NA
#' #>
#' @noRd
stri_locate_all_regex <- function(str, pattern) {
  str <- enc2utf8(as.character(str))
  lapply(gregexpr(pcre_pattern(pattern), str, perl = TRUE), function(m) {
    if (is.na(m[1L]) || m[1L] == -1L) {
      return(matrix(NA_integer_, 1L, 2L, dimnames = list(NULL, c("start", "end"))))
    }
    len <- attr(m, "match.length")
    cbind(start = as.integer(m), end = as.integer(m) + len - 1L)
  })
}

#' Locate the last occurrence of a fixed string
#'
#' Replacement for `stringi::stri_locate_last_fixed()`, written in C++. Returns
#' a matrix with the columns `start` and `end` in characters.
#' @param str character vector
#' @param pattern single non empty string
#' @examples
#' stri_locate_last_fixed(c("a</x></x>", "abc"), "</x>")
#' #>      start end
#' #> [1,]     6   9
#' #> [2,]    NA  NA
#' @noRd
stri_locate_last_fixed <- function(str, pattern) {
  res <- stri_locate_last_fixed_cpp(as.character(str), as.character(pattern))
  dimnames(res) <- list(NULL, c("start", "end"))
  res
}

#' Split at a fixed string
#'
#' Replacement for `stringi::stri_split_fixed()`. `strsplit()` drops a trailing
#' empty piece and returns `character(0)` for "", stringi keeps both. This is
#' corrected here. No `n` or `omit_empty` arguments.
#' @param str character vector
#' @param pattern single string
#' @examples
#' stri_split_fixed(c("A1:B2", "A1:", ""), ":")
#' #> [[1]]
#' #> [1] "A1" "B2"
#' #>
#' #> [[2]]
#' #> [1] "A1" ""
#' #>
#' #> [[3]]
#' #> [1] ""
#' #>
#' @noRd
stri_split_fixed <- function(str, pattern) {
  str <- enc2utf8(as.character(str))
  res <- strsplit(str, enc2utf8(pattern), fixed = TRUE)

  # strsplit() drops a trailing empty string and returns character(0) for ""
  empty <- which(!is.na(str) & !nzchar(str))
  res[empty] <- list("")
  tail_empty <- which(!is.na(str) & nzchar(str) & endsWith(str, pattern))
  res[tail_empty] <- lapply(res[tail_empty], c, "")
  res[is.na(str)] <- list(NA_character_)

  res
}

#' Substrings
#'
#' Replacement for `stringi::stri_sub()` with `from` and `to`. Negative
#' positions count from the end. There is no `length` argument.
#' @param str character vector
#' @param from first position
#' @param to last position
#' @examples
#' stri_sub("abcdef", 2, 4)
#' #> [1] "bcd"
#' stri_sub("abcdef", -3, -2)
#' #> [1] "de"
#' stri_sub("abcdef", 4, 100)
#' #> [1] "def"
#' @noRd
stri_sub <- function(str, from = 1L, to = -1L) {
  str <- enc2utf8(as.character(str))

  if (length(from) == 1L && length(to) == 1L && !is.na(from) && !is.na(to) && from >= 1L && to >= 0L) {
    return(substr(str, from, to))
  }

  n <- nchar(str)
  from <- rep_len(from, length(str))
  to <- rep_len(to, length(str))

  # negative positions count from the end
  neg <- !is.na(from) & from < 0L
  from[neg] <- n[neg] + from[neg] + 1L
  neg <- !is.na(to) & to < 0L
  to[neg] <- n[neg] + to[neg] + 1L

  substring(str, pmax(from, 1L), to)
}

#' Replace a part of a string
#'
#' Replacement for `stringi::stri_sub_replace()` with `from`, `to` and `value`.
#' `to = from - 1` inserts `value` in front of `from`. The C++ version is used
#' for positive positions, anything else (negative positions, `NA`) goes
#' through a slower R version. Both recycle `from`, `to` and `value` along
#' `str`, but not `str` along longer positions.
#' @inheritParams stri_sub
#' @param value replacement
#' @examples
#' stri_sub_replace("abcdef", 2, 3, value = "XYZ")
#' #> [1] "aXYZdef"
#' stri_sub_replace("abcdef", 3, 2, value = "-")
#' #> [1] "ab-cdef"
#' @noRd
stri_sub_replace <- function(str, from = 1L, to = -1L, value) {
  str <- as.character(str)
  value <- as.character(value)

  # `to = from - 1L` inserts value in front of from
  if (length(str) && length(from) && length(to) && length(value) &&
      !anyNA(from) && !anyNA(to) && all(from >= 1L) && all(to >= from - 1L)) {
    return(stri_sub_replace_cpp(str, as.integer(from), as.integer(to), value))
  }

  str <- enc2utf8(str)
  value <- enc2utf8(value)
  n <- nchar(str)
  from <- rep_len(from, length(str))
  to <- rep_len(to, length(str))

  # negative positions count from the end
  neg <- !is.na(from) & from < 0L
  from[neg] <- n[neg] + from[neg] + 1L
  neg <- !is.na(to) & to < 0L
  to[neg] <- n[neg] + to[neg] + 1L
  from <- pmax(from, 1L)
  to <- pmax(pmin(to, n), from - 1L)

  stri_join(
    substr(str, 1L, from - 1L),
    value,
    substr(str, to + 1L, n)
  )
}

#' Replace several parts of one string
#'
#' Replacement for `stringi::stri_sub_replace_all()`. Works on a single string.
#' The ranges must be sorted and must not overlap, stringi does not need this.
#' @param str single string
#' @param from,to start and end of the ranges
#' @param replacement replacements, recycled along the ranges
#' @examples
#' stri_sub_replace_all("a1b22c", from = c(2, 4), to = c(2, 5), replacement = c("X", "Y"))
#' #> [1] "aXbYc"
#' @noRd
stri_sub_replace_all <- function(str, from, to, replacement) {
  str <- enc2utf8(as.character(str))
  if (!length(from)) return(str)

  # the ranges are sorted and do not overlap
  replacement <- rep_len(enc2utf8(as.character(replacement)), length(from))
  keep <- substring(str, c(1L, to + 1L), c(from - 1L, nchar(str)))

  paste(c(rbind(keep, c(replacement, ""))), collapse = "")
}

#' Pad strings on the left
#'
#' Replacement for `stringi::stri_pad_left()`, written in C++. `pad` has to be a
#' single code point. stringi measures the display width of a string: combining
#' marks, the soft hyphen and control characters count as 0, some symbols and
#' emoji as 2. This counts code points, which is the same for all other
#' characters.
#' @param str character vector
#' @param width minimal width
#' @param pad single character
#' @examples
#' stri_pad_left(c("7", "42", "12345"), 4, "0")
#' #> [1] "0007"  "0042"  "12345"
#' stri_pad_left("1", 3, "\u00e4") == "\u00e4\u00e41"
#' #> [1] TRUE
#' @noRd
stri_pad_left <- function(str, width = floor(0.9 * getOption("width")), pad = " ") {
  if (!length(str) || !length(width) || !length(pad)) return(character())
  stri_pad_left_cpp(as.character(str), as.integer(width), as.character(pad))
}

#' Unique strings
#'
#' Replacement for `stringi::stri_unique()`, written in C++. stringi compares
#' strings by canonical equivalence and ignores some control characters, so it
#' can merge strings that differ in bytes. This compares the bytes. Order of
#' first occurrence is kept. Used for shared strings, where identical bytes are
#' what matters.
#' @param str character vector
#' @examples
#' stri_unique(c("b", "a", "b", NA, NA, "a"))
#' #> [1] "b" "a" NA
#' @noRd
stri_unique <- function(str) {
  stri_unique_cpp(as.character(str))
}

#' Random strings
#'
#' Replacement for `stringi::stri_rand_strings()`. With the same seed it returns
#' the same strings as stringi, because it draws one number from R's uniform
#' generator per character, like stringi does. For this the "Rounding" sampler
#' is used temporarily (R >= 3.6.0). Only printable ASCII characters can be
#' selected with `pattern`, stringi accepts any Unicode character class. Only a
#' single `length` is supported.
#' @param n number of strings
#' @param length number of characters per string
#' @param pattern character class (PCRE) the characters are taken from
#' @examples
#' set.seed(1); x <- stri_rand_strings(2, 8, "[A-F0-9]")
#' nchar(x)
#' #> [1] 8 8
#' all(grepl("^[A-F0-9]{8}$", x))
#' #> [1] TRUE
#' @noRd
stri_rand_strings <- function(n, length, pattern = "[A-Za-z0-9]") {
  # the candidates are all printable ASCII characters that match the pattern
  chars <- intToUtf8(32:126, multiple = TRUE)
  chars <- chars[grepl(pattern, chars, perl = TRUE)]

  if (length(n) != 1L || n < 1L) return(character())

  # like stringi: one draw of the uniform generator per character, i.e.
  # floor(unif_rand() * N). this is what sample.int() does with the "Rounding"
  # sampler (the only one before R 3.6.0)
  if (getRversion() >= "3.6.0") {
    old <- suppressWarnings(RNGkind(sample.kind = "Rounding"))[3L]
    on.exit(suppressWarnings(RNGkind(sample.kind = old)))
  }
  idx <- suppressWarnings(sample.int(length(chars), n * length, replace = TRUE))
  res <- matrix(chars[idx], nrow = n, ncol = length, byrow = TRUE)
  do.call(paste0, as.data.frame(res, stringsAsFactors = FALSE))
}

#' Read lines of a file
#'
#' Replacement for `stringi::stri_read_lines()`. Uses `readLines()` and removes
#' a byte order mark from the first line. An empty file returns "".
#' @param con file name or connection
#' @param encoding encoding of the file
#' @examples
#' f <- tempfile()
#' writeBin(c(as.raw(c(0xEF, 0xBB, 0xBF)), charToRaw("a\nb\n")), f)
#' stri_read_lines(f)
#' #> [1] "a" "b"
#' @noRd
stri_read_lines <- function(con, encoding = "UTF-8") {
  res <- readLines(con, encoding = encoding, warn = FALSE)
  if (!length(res)) return("")

  # like stringi, remove a byte order mark
  bom <- as.raw(c(0xEF, 0xBB, 0xBF))
  first <- charToRaw(res[1L])
  if (length(first) >= 3L && identical(first[1:3], bom)) {
    res[1L] <- rawToChar(first[-(1:3)])
    Encoding(res[1L]) <- "UTF-8"
  }

  res
}

#' Escape non ASCII characters
#'
#' Replacement for `stringi::stri_escape_unicode()`. Everything that is not
#' printable ASCII is written as `\uXXXX` (`\UXXXXXXXX` above U+FFFF), as are
#' quotes and backslashes. Control characters like tab and newline get their
#' short form.
#' @param str character vector
#' @examples
#' stri_escape_unicode("a\u00e4\u20ac\n")
#' #> [1] "a\\u00e4\\u20ac\\n"
#' @noRd
stri_escape_unicode <- function(str) {
  str <- enc2utf8(as.character(str))
  special <- c("\\" = "\\\\", "\"" = "\\\"", "'" = "\\'", "\a" = "\\a", "\b" = "\\b",
               "\t" = "\\t", "\n" = "\\n", "\v" = "\\v", "\f" = "\\f", "\r" = "\\r")

  esc <- function(chars) {
    vapply(chars, function(ch) {
      if (!is.na(sp <- special[ch])) return(sp)
      cp <- utf8ToInt(ch)
      if (cp > 0xFFFFL) sprintf("\\U%08x", cp) else sprintf("\\u%04x", cp)
    }, NA_character_, USE.NAMES = FALSE)
  }

  pat <- "[^ -~]|[\\\\\"']"
  m <- gregexpr(pat, str, perl = TRUE)
  na <- is.na(str)
  regmatches(str, m) <- lapply(regmatches(str, m), esc)
  str[na] <- NA_character_
  str
}

#' Resolve Unicode escapes
#'
#' Replacement for `stringi::stri_unescape_unicode()`. Understands `\uXXXX`
#' (surrogate pairs are combined), `\UXXXXXXXX`, `\xHH`, `\x{H...}`, octal
#' escapes, `\cX` control characters and the single character escapes.
#' Incomplete escapes and code points above U+10FFFF give `NA`, an escape for
#' the NUL character raises an error, as in stringi.
#' @param str character vector
#' @examples
#' utf8ToInt(stri_unescape_unicode("a\\u00e4\\n"))
#' #> [1]  97 228  10
#' utf8ToInt(stri_unescape_unicode("\\u20ac"))
#' #> [1] 8364
#' stri_unescape_unicode("\\u00e")
#' #> [1] NA
#' @noRd
stri_unescape_unicode <- function(str) {
  str <- enc2utf8(as.character(str))

  # \uXXXX (surrogate pairs are combined), \UXXXXXXXX, \xHH, \x{H..}, \cX,
  # octal and the single character escapes. other escaped characters stay as
  # they are
  pat <- paste0(
    "(?s)\\\\(?:",
    "u[dD][89abAB][0-9a-fA-F]{2}\\\\u[dD][c-fC-F][0-9a-fA-F]{2}|",
    "u[0-9a-fA-F]{4}|U[0-9a-fA-F]{8}|x\\{[0-9a-fA-F]+\\}|x[0-9a-fA-F]{1,2}|",
    "c[\\s\\S]|[0-7]{1,3}|.)"
  )
  single <- c(a = "\a", b = "\b", t = "\t", n = "\n", v = "\v", f = "\f", r = "\r")

  # decodes one escape. NA for an incomplete escape or a code point above
  # U+10FFFF, a longer string for escapes that raise an error
  dec <- function(tok) {
    body <- substring(tok, 2L)
    type <- substr(body, 1L, 1L)
    if (nchar(body) == 1L && type %in% c("u", "U", "x")) return(NA_character_)
    if (nchar(body) == 1L && !grepl("[0-7]", body)) {
      if (!is.na(sp <- single[body])) return(unname(sp))
      return(body)
    }
    cp <- if (type == "c") {
      bitwAnd(utf8ToInt(substring(body, 2L)), 31L)
    } else if (type == "u" && nchar(body) > 5L) {
      hi <- strtoi(substr(body, 2L, 5L), 16L)
      lo <- strtoi(substr(body, 8L, 11L), 16L)
      0x10000L + (hi - 0xD800L) * 0x400L + (lo - 0xDC00L)
    } else if (type == "u" || type == "U") {
      strtoi(substring(body, 2L), 16L)
    } else if (type == "x") {
      strtoi(gsub("[{}]", "", substring(body, 2L)), 16L)
    } else {
      strtoi(body, 8L)
    }
    if (is.na(cp) || cp > 0x10FFFFL) return(NA_character_)
    if (cp == 0L || (cp >= 0xD800L && cp <= 0xDFFFL)) return("error")
    intToUtf8(cp)
  }

  res <- str
  for (i in which(!is.na(str))) {
    m <- gregexpr(pat, str[i], perl = TRUE)[[1L]]
    found <- m[1L] != -1L
    decoded <- if (found) vapply(regmatches(str[i], list(m))[[1L]], dec, NA_character_, USE.NAMES = FALSE) else character()

    # a backslash at the end that is not part of an escape is invalid
    last_end <- if (found) max(m + attr(m, "match.length")) - 1L else 0L
    if (anyNA(decoded) || (endsWith(str[i], "\\") && last_end < nchar(str[i]))) {
      res[i] <- NA_character_
    } else if (any(nchar(decoded) > 1L)) {
      stop("invalid or unsupported escape sequence", call. = FALSE)
    } else if (found) {
      regmatches(res[i], list(m)) <- list(decoded)
    }
  }
  res
}

#' Collator options
#'
#' Replacement for `stringi::stri_opts_collator()`. stringi hands these options
#' to ICU. Only `numeric` is supported here.
#' @param numeric compare digit sequences by value
#' @examples
#' stri_opts_collator(numeric = TRUE)
#' #> $numeric
#' #> [1] TRUE
#' #>
#' @noRd
stri_opts_collator <- function(numeric = FALSE) {
  list(numeric = numeric)
}

#' Order strings
#'
#' Replacement for `stringi::stri_order()`. stringi sorts by the collation rules
#' of the locale (for example "a" before "B"). This sorts by code points and
#' does not depend on the locale, so uppercase letters come before lowercase
#' ones. With `numeric = TRUE` digit sequences are compared by value, as in
#' stringi. Used to order part names like `sheet2.xml` and
#' `sheet10.xml`.
#' @param str character vector
#' @param opts_collator list returned by the collator options
#' @examples
#' x <- c("sheet10.xml", "sheet2.xml", "sheet1.xml")
#' x[stri_order(x, stri_opts_collator(numeric = TRUE))]
#' #> [1] "sheet1.xml"  "sheet2.xml"  "sheet10.xml"
#' y <- c("b", "B", "a")
#' y[stri_order(y)]
#' #> [1] "B" "a" "b"
#' @noRd
stri_order <- function(str, opts_collator = stri_opts_collator()) {
  str <- enc2utf8(as.character(str))
  if (isTRUE(opts_collator$numeric)) return(stri_order_numeric_cpp(str))
  order(str, method = "radix")
}
