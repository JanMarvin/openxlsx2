
createFillNode <- function(style, patternType = "solid") {
  # TODO assert_class(style, "Style")
  fill <- style$fill

  ## gradientFill
  if (any(grepl("gradientFill", fill))) {
    fillNode <- fill # stri_join("<fill>", fill, "</fill>")
  } else if (!is.null(fill$fillFg) | !is.null(fill$fillBg)) {
    fillNode <-
      stri_join(
        "<fill>",
        sprintf('<patternFill patternType="%s">', patternType)
      )

    if (!is.null(fill$fillFg)) {
      fillNode <-
        stri_join(fillNode, sprintf(
          "<fgColor %s/>",
          stri_join(
            stri_join(names(fill$fillFg), '="', fill$fillFg, '"'),
            sep = " ",
            collapse = " "
          )
        ))
    }

    if (!is.null(fill$fillBg)) {
      fillNode <-
        stri_join(fillNode, sprintf(
          "<bgColor %s/>",
          stri_join(
            stri_join(names(fill$fillBg), '="', fill$fillBg, '"'),
            sep = " ",
            collapse = " "
          )
        ))
    }

    fillNode <- stri_join(fillNode, "</patternFill></fill>")
  } else {
    return(NULL)
  }

  return(fillNode)
}

createBorderNode <- function(style) {
  # TODO assert_class(style, "Style")
  borderNode <- "<border"

  if (style$borderDiagonalUp) {
    borderNode <- stri_join(borderNode, 'diagonalUp="1"', sep = " ")
  }

  if (style$borderDiagonalDown) {
    borderNode <-
      stri_join(borderNode, 'diagonalDown="1"', sep = " ")
  }

  borderNode <- stri_join(borderNode, ">")

  if (!is.null(style$borderLeft)) {
    borderNode <-
      stri_join(
        borderNode,
        sprintf('<left style="%s">', style$borderLeft),
        sprintf(
          '<color %s="%s"/>',
          names(style$borderLeftColour),
          style$borderLeftColour
        ),
        "</left>"
      )
  }

  if (!is.null(style$borderRight)) {
    borderNode <-
      stri_join(
        borderNode,
        sprintf('<right style="%s">', style$borderRight),
        sprintf(
          '<color %s="%s"/>',
          names(style$borderRightColour),
          style$borderRightColour
        ),
        "</right>"
      )
  }

  if (!is.null(style$borderTop)) {
    borderNode <-
      stri_join(
        borderNode,
        sprintf('<top style="%s">', style$borderTop),
        sprintf(
          '<color %s="%s"/>',
          names(style$borderTopColour),
          style$borderTopColour
        ),
        "</top>"
      )
  }

  if (!is.null(style$borderBottom)) {
    borderNode <-
      stri_join(
        borderNode,
        sprintf('<bottom style="%s">', style$borderBottom),
        sprintf(
          '<color %s="%s"/>',
          names(style$borderBottomColour),
          style$borderBottomColour
        ),
        "</bottom>"
      )
  }

  if (!is.null(style$borderDiagonal)) {
    borderNode <-
      stri_join(
        borderNode,
        sprintf('<diagonal style="%s">', style$borderDiagonal),
        sprintf(
          '<color %s="%s"/>',
          names(style$borderDiagonalColour),
          style$borderDiagonalColour
        ),
        "</diagonal>"
      )
  }

  stri_join(borderNode, "</border>")
}
