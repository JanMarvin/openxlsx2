
#' Clean worksheet name
#'
#' @param x A vector, coherced to `character`
#' @returns x with bad characters removed
#' @export
clean_worksheet_name <- function(x) {
  replaceIllegalCharacters(x)
}

#' Detect illegal characters
#' @param x A vector, coerced to character
#' @returns A `logical` vector
#' @noRd
any_illegal_chars <- function(x) {
  x <- as.character(x)
  any(vapply(illegal_chars(), grepl, logical(length(x)), x = x, fixed = TRUE))
}


illegal_chars <- function() { c("&",     '"',      "'",      "<",    ">",    "\a", "\b", "\v", "\f") } # nolint
legal_chars   <- function() { c("&amp;", "&quot;", "&apos;", "&lt;", "&gt;", "",   "",   "",   "") }   # nolint

#' converts & to &amp;
#' @param v some xml string
#' @keywords internal
#' @noRd
replaceIllegalCharacters <- function(v) {
  v <- as.character(v)
  bad <- Encoding(v) != "UTF-8"

  if (any(bad)) {
    v[bad] <- stringi::stri_conv(v[bad], from = "", to = "UTF-8")
  }

  stringi::stri_replace_all_fixed(v, illegal_chars(), legal_chars(), vectorize_all = FALSE)
}

#' converts &amp; to &
#' @param v some xml string
#' @keywords internal
#' @noRd
replaceXMLEntities <- function(v) {
  v <- gsub("&amp;",  "&", v, fixed = TRUE)
  v <- gsub("&quot;", '"', v, fixed = TRUE)
  v <- gsub("&apos;", "'", v, fixed = TRUE)
  v <- gsub("&lt;",   "<", v, fixed = TRUE)
  v <- gsub("&gt;",   ">", v, fixed = TRUE)
  return(v)
}
