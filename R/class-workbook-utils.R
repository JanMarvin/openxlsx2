#' Validate sheet
#'
#' @param wb A workbook
#' @param sheetName The sheet name to validate
#' @return The sheet name -- or the position?  This should be consistent
#' @noRd
wb_validate_sheet <- function(wb, sheetName) {
  assert_workbook(wb)

  if (!is.numeric(sheetName)) {
    if (is.null(wb$sheet_names)) {
      stop("wb does not contain any worksheets.", call. = FALSE)
    }
  }

  if (is.numeric(sheetName)) {
    if (sheetName > length(wb$sheet_names)) {
      msg <- sprintf("wb only contains %i sheets.", length(wb$sheet_names))
      stop(msg, call. = FALSE)
    }

    return(sheetName)
  }

  if (!sheetName %in% replaceXMLEntities(wb$sheet_names)) {
    msg <- sprintf("Sheet '%s' does not exist.", replaceXMLEntities(sheetName))
    stop(msg, call. = FALSE)
  }


  which(replaceXMLEntities(wb$sheet_names) == sheetName)
}


#' Create a font node from a style
#'
#' @param wb a workbook
#' @param style style
#' @return The font node as xml?
#' @noRd
wb_create_font_node <- function(wb, style) {
  assert_style(wb)
  # assert_style(style)

  baseFont <- wb$getBaseFont()

  # TODO implement paste_c()
  fontNode <- "<font>"

  ## size
  if (is.null(style$fontSize[[1]])) {
    fontNode <- stri_join(fontNode, sprintf('<sz %s="%s"/>', names(baseFont$size), baseFont$size))
  } else {
    fontNode <- stri_join(fontNode, sprintf('<sz %s="%s"/>', names(style$fontSize), style$fontSize))
  }

  ## colour
  if (is.null(style$fontColour[[1]])) {
    fontNode <- stri_join(
      fontNode,
      sprintf('<color %s="%s"/>', names(baseFont$colour), baseFont$colour)
    )
  } else {
    if (length(style$fontColour) > 1) {
      fontNode <- stri_join(
        fontNode,
        sprintf(
          "<color %s/>",
          stri_join(
            sapply(
              seq_along(style$fontColour),
              function(i) {
                sprintf('%s="%s"', names(style$fontColour)[i], style$fontColour[i])
              }
            ),
            sep = " ",
            collapse = " "
          )
        )
      )
    } else {
      fontNode <- stri_join(
        fontNode,
        sprintf('<color %s="%s"/>', names(style$fontColour), style$fontColour)
      )
    }
  }


  ## name
  if (is.null(style$fontName[[1]])) {
    fontNode <- stri_join(
      fontNode,
      sprintf('<name %s="%s"/>', names(baseFont$name), baseFont$name)
    )
  } else {
    fontNode <- stri_join(
      fontNode,
      sprintf('<name %s="%s"/>', names(style$fontName), style$fontName)
    )
  }

  ### Create new font and return Id
  if (!is.null(style$fontFamily)) {
    fontNode <- stri_join(fontNode, sprintf('<family val = "%s"/>', style$fontFamily))
  }

  if (!is.null(style$fontScheme)) {
    fontNode <- stri_join(fontNode, sprintf('<scheme val = "%s"/>', style$fontScheme))
  }

  if ("BOLD" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, "<b/>")
  }

  if ("ITALIC" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, "<i/>")
  }

  if ("UNDERLINE" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, '<u val="single"/>')
  }

  if ("UNDERLINE2" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, '<u val="double"/>')
  }

  if ("STRIKEOUT" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, "<strike/>")
  }

  stri_join(fontNode, "</font>")
}
