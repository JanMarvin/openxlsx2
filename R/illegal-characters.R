
#' Clean worksheet name
#'
#' @param x A vector, coherced to `character`
#' @returns x with bad characters removed
#' @export
clean_worksheet_name <- function(x) {
  replace_legal_chars(x)
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
#' @keywords internal
#' @noRd
replace_legal_chars <- function(x) {
  x <- as.character(x)
  bad <- Encoding(x) != "UTF-8"

  if (any(bad)) {
    x[bad] <- stringi::stri_conv(x[bad], from = "", to = "UTF-8")
  }

  stringi::stri_replace_all_fixed(x, legal_chars(), legal_sub(), vectorize_all = FALSE)
}

replaceIllegalCharacters <- replace_legal_chars

#' converts &amp; to &
#' @param x some xml string
#' @keywords internal
#' @noRd
replaceXMLEntities <- function(x) {
  stringi::stri_replace_all_fixed(
    x,
    c("&amp;", "&quot;", "&apos;", "&lt;", "&gt;"),
    c("&",     '"',      "'",      "<",    ">"),
    vectorize_all = FALSE
  )
}

