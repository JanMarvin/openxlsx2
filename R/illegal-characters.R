#' Detect illegal characters
#' @param x A vector, coerced to character
#' @returns A `logical` vector
#' @noRd
has_illegal_chars <- function(x) {
  res <- vapply(
    illegal_chars(),
    function(i) stringi::stri_detect_fixed(x, i),
    logical(length(x)),
    USE.NAMES = FALSE
  )

  if (is.null(dim(res))) {
    any(res)
  } else {
    as.logical(apply(res, 1, any))
  }
}

# vectors for character types
illegal_chars <- function() { c("\\", "/", "?", "*", ":", "[", "]") } # nolint
legal_chars   <- function() { c("&",     '"',      "'",      "<",    ">",    "\a", "\b", "\v", "\f") } # nolint
legal_sub     <- function() { c("&amp;", "&quot;", "&apos;", "&lt;", "&gt;", "",   "",   "",   ""  ) } # nolint

#' remove characters that are not allowed in XML 1.0
#'
#' Removes raw control characters (U+0000 to U+001F except tab, line feed and
#' carriage return), the noncharacters U+FFFE and U+FFFF, and character
#' references to them like `&#01;` or `&#x1;`. None of them are allowed in
#' XML 1.0 and Excel refuses to open files containing them.
#' @param x some xml string
#' @noRd
strip_xml_control_chars <- function(x) {
  # most strings contain none of these, a byte wise search is much faster than the replacement
  # U+FFFE and U+FFFF are the UTF-8 byte sequences EF BF BE and EF BF BF
  hit <- grepl("[\\x{01}-\\x{08}\\x{0B}\\x{0C}\\x{0E}-\\x{1F}]|\\xEF\\xBF[\\xBE\\xBF]|&#", x, perl = TRUE, useBytes = TRUE)
  if (any(hit)) {
    x[hit] <- stringi::stri_replace_all_regex(
      x[hit],
      paste0(
        "[\\x{01}-\\x{08}\\x{0B}\\x{0C}\\x{0E}-\\x{1F}\\x{FFFE}\\x{FFFF}]",
        "|&#(?:[xX]0*(?:[0-8bBcCeEfF]|1[0-9a-fA-F]|[fF]{3}[eEfF])|0*(?:[0-8]|1[124-9]|2[0-9]|3[01]|6553[45]));"
      ),
      ""
    )
  }
  x
}

#' converts & to &amp;
#' @param x some xml string
#' @noRd
replace_legal_chars <- function(x) {
  x <- as.character(x)
  stringi::stri_replace_all_fixed(x, legal_chars(), legal_sub(), vectorize_all = FALSE)
}

replace_illegal_chars <- function(x, replacement = " ") {
  x <- as.character(x)
  x <- stringi::stri_escape_unicode(x)
  stringi::stri_replace_all_fixed(x, illegal_chars(), replacement, vectorize_all = FALSE)
}

# TODO investigate if this should be integrated into replace_legal_chars()
escape_newline_and_tab <- function(x) {
  stringi::stri_replace_all_fixed(
    x,
    c("\n",       "\t"),
    c("_x000a_", "_x0009_"),
    vectorize_all = FALSE
  )
}

#' Clean worksheet name
#'
#' Cleans a worksheet name by removing legal characters.
#'
#' @details Illegal characters are considered `\`, `/`, `?`, `*`, `:`, `[`, and
#' `]`.  These must be intentionally removed from worksheet names prior to
#' creating a new worksheet.
#'
#' @param x A vector, coerced to `character`
#' @param replacement A single value to replace illegal characters by.
#' @returns x with bad characters removed
clean_worksheet_name <- function(x, replacement = " ") {
  stopifnot(length(replacement) == 1, !has_illegal_chars(replacement))
  replace_illegal_chars(x, replacement = replacement)
}

#' converts &amp; to &
#' @param x some xml string
#' @noRd
replaceXMLEntities <- function(x) {
  stringi::stri_replace_all_fixed(
    x,
    c("&amp;", "&quot;", "&apos;", "&lt;", "&gt;"),
    c("&",     '"',      "'",      "<",    ">"),
    vectorize_all = FALSE
  )
}

# for dateFormats
escape_format_code <- function(formatCode) {

  formatCode <- replace_legal_chars(replaceXMLEntities(formatCode))

  # Check if it's a date/time format
  is_date_time <- grepl("(y{2,4}|m{1,4}|d{1,2}|h{1,2}|s{1,2})", tolower(formatCode))

  if (is_date_time && grepl("/", formatCode)) {
    # Use a placeholder for AM/PM
    has_ampm <- grepl("AM/PM", formatCode, ignore.case = TRUE)
    has_ap   <- grepl("A/P", formatCode, ignore.case = TRUE)

    if (has_ampm) formatCode <- gsub("AM/PM", "@@AMPM@@", formatCode, ignore.case = TRUE)
    if (has_ap)   formatCode <- gsub("A/P", "@@AP@@", formatCode, ignore.case = TRUE)

    formatCode <- gsub("(?<!\\\\)/", "\\\\/", formatCode, perl = TRUE)

    if (has_ampm) formatCode <- gsub("@@AMPM@@", "AM/PM", formatCode)
    if (has_ap)   formatCode <- gsub("@@AP@@", "A/P", formatCode)
  }

  formatCode
}
