
# TODO rename file or move somewhere else?

get_style_max_char_width <- function(thisStyle) {
  # TODO assert_class() style
  fN <- unlist(thisStyle$fontName, use.names = FALSE)
  if (is.null(fN)) {
    fN <- "calibri"
  } else {
    fN <- gsub(" ", ".", tolower(fN), fixed = TRUE)
    if (!fN %in% names(openxlsxFontSizeLookupTable)) {
      fN <- "calibri"
    }
  }

  fS <- unlist(thisStyle$fontSize, use.names = FALSE)
  if (is.null(fS)) {
    fS <- 11
  } else {
    fS <- as.numeric(fS)
    if (fS < 8) {
      fS <- 8
    } else if (fS > 36) {
      fS <- 36
    }
  }

  if ("BOLD" %in% thisStyle$fontDecoration) {
    styleMaxCharWidth <- openxlsxFontSizeLookupTableBold[[fN]][fS - 7]
  } else {
    styleMaxCharWidth <- openxlsxFontSizeLookupTable[[fN]][fS - 7]
  }

  return(styleMaxCharWidth)
}
