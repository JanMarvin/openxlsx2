
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

#' converts & to &amp;
#' @param x some xml string
#' @noRd
replace_legal_chars <- function(x) {
  x <- as.character(x)
  bad <- Encoding(x) != "UTF-8"

  if (any(bad)) {
    x[bad] <- stringi::stri_conv(x[bad], from = "", to = "UTF-8")
  }

  stringi::stri_replace_all_fixed(x, legal_chars(), legal_sub(), vectorize_all = FALSE)
}

replace_illegal_chars <- function(x, replacement = " ") {
  x <- as.character(x)
  x <- stringi::stri_escape_unicode(x)
  stringi::stri_replace_all_fixed(x, illegal_chars(), replacement, vectorize_all = FALSE)
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
