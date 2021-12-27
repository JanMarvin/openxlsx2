
# class -------------------------------------------------------------------

Style <- setRefClass(
  "Style",
  fields = c(
    "fontId",
    "fontName",
    "fontColour",
    "fontSize",
    "fontFamily",
    "fontScheme",
    "fontDecoration",
    "borderTop",
    "borderLeft",
    "borderRight",
    "borderBottom",
    "borderTopColour",
    "borderLeftColour",
    "borderRightColour",
    "borderBottomColour",
    "borderDiagonal",
    "borderDiagonalColour",
    "borderDiagonalUp",
    "borderDiagonalDown",
    "halign",
    "valign",
    "indent",
    "textRotation",
    "numFmt",
    "fill",
    "wrapText",
    "locked",
    "hidden",
    "xfId"
  ),
  methods = list(
    initialize = function() {
      .self$fontId <- NULL
      .self$fontName <- NULL
      .self$fontColour <- NULL
      .self$fontSize <- NULL
      .self$fontFamily <- NULL
      .self$fontScheme <- NULL
      .self$fontDecoration <- NULL

      .self$borderTop <- NULL
      .self$borderLeft <- NULL
      .self$borderRight <- NULL
      .self$borderBottom <- NULL
      .self$borderTopColour <- NULL
      .self$borderLeftColour <- NULL
      .self$borderRightColour <- NULL
      .self$borderBottomColour <- NULL
      .self$borderDiagonal <- NULL
      .self$borderDiagonalColour <- NULL
      .self$borderDiagonalUp <- FALSE
      .self$borderDiagonalDown <- FALSE

      .self$halign <- NULL
      .self$valign <- NULL
      .self$indent <- NULL
      .self$textRotation <- NULL
      .self$numFmt <- NULL
      .self$fill <- NULL
      .self$wrapText <- NULL
      .self$hidden <- NULL
      .self$locked <- NULL
      .self$xfId <- NULL
    },

    show = function(print = TRUE) {
      numFmtMapping <- list(
        list("numFmtId" = 0),
        list("numFmtId" = 2),
        list("numFmtId" = 164),
        list("numFmtId" = 44),
        list("numFmtId" = 14),
        list("numFmtId" = 167),
        list("numFmtId" = 10),
        list("numFmtId" = 11),
        list("numFmtId" = 49)
      )

      validNumFmt <- c("GENERAL", "NUMBER", "CURRENCY", "ACCOUNTING", "DATE", "TIME", "PERCENTAGE", "SCIENTIFIC", "TEXT")

      if (!is.null(numFmt)) {
        if (as.integer(numFmt$numFmtId) %in% unlist(numFmtMapping)) {
          numFmtStr <- validNumFmt[unlist(numFmtMapping) == as.integer(numFmt$numFmtId)]
        } else {
          numFmtStr <- sprintf('"%s"', numFmt$formatCode)
        }
      } else {
        numFmtStr <- "GENERAL"
      }

      borders <- c(sprintf("Top: %s", borderTop), sprintf("Bottom: %s", borderBottom), sprintf("Left: %s", borderLeft), sprintf("Right: %s", borderRight))
      borderColours <- gsub("^FF", "#", c(borderTopColour, borderBottomColour, borderLeftColour, borderRightColour))

      fgFill <- fill$fillFg
      bgFill <- fill$fillBg

      styleShow <- "A custom cell style. \n\n"

      styleShow <- append(styleShow, sprintf("Cell formatting: %s \n", numFmtStr)) ## numFmt
      styleShow <- append(styleShow, sprintf("Font name: %s \n", fontName[[1]])) ## Font name
      styleShow <- append(styleShow, sprintf("Font size: %s \n", fontSize[[1]])) ## Font size
      styleShow <- append(styleShow, sprintf("Font colour: %s \n", gsub("^FF", "#", fontColour[[1]]))) ## Font colour

      ## Font decoration
      if (length(fontDecoration) > 0) {
        styleShow <- append(styleShow, sprintf("Font decoration: %s \n", paste(fontDecoration, collapse = ", ")))
      }

      if (length(borders) > 0) {
        styleShow <- append(styleShow, sprintf("Cell borders: %s \n", paste(borders, collapse = ", "))) ## Cell borders
        styleShow <- append(styleShow, sprintf("Cell border colours: %s \n", paste(borderColours, collapse = ", "))) ## Cell borders
      }

      if (!is.null(halign)) {
        styleShow <- append(styleShow, sprintf("Cell horz. align: %s \n", halign))
      } ## Cell horizontal alignment

      if (!is.null(valign)) {
        styleShow <- append(styleShow, sprintf("Cell vert. align: %s \n", valign))
      } ## Cell vertical alignment

      if (!is.null(indent)) {
        styleShow <- append(styleShow, sprintf("Cell indent: %s \n", indent))
      } ## Cell indent

      if (!is.null(textRotation)) {
        styleShow <- append(styleShow, sprintf("Cell text rotation: %s \n", textRotation))
      } ## Cell text rotation

      ## Cell fill colour
      if (length(fgFill) > 0) {
        styleShow <- append(styleShow, sprintf("Cell fill foreground: %s \n", paste(paste0(names(fgFill), ": ", sub("^FF", "#", fgFill)), collapse = ", ")))
      }

      if (length(bgFill) > 0) {
        styleShow <- append(styleShow, sprintf("Cell fill background: %s \n", paste(paste0(names(bgFill), ": ", sub("^FF", "#", bgFill)), collapse = ", ")))
      }

      if (!is.null(locked)) {
        styleShow <- append(styleShow, sprintf("Cell protection: %s \n", locked))
      } ## Cell protection
      if (!is.null(hidden)) {
        styleShow <- append(styleShow, sprintf("Cell formula hidden: %s \n", hidden))
      } ## Cell formula hidden

      styleShow <- append(styleShow, sprintf("wraptext: %s", wrapText)) ## wrap text

      styleShow <- c(styleShow, "\n\n")

      if (print) {
        cat(styleShow)
      }

      return(invisible(styleShow))
    },

    as.list = function() {
      l <- list(
        "fontId" = fontId,
        "fontName" = fontName,
        "fontColour" = fontColour,
        "fontSize" = fontSize,
        "fontFamily" = fontFamily,
        "fontScheme" = fontScheme,
        "fontDecoration" = fontDecoration,

        "borderTop" = borderTop,
        "borderLeft" = borderLeft,
        "borderRight" = borderRight,
        "borderBottom" = borderBottom,
        "borderTopColour" = borderTopColour,
        "borderLeftColour" = borderLeftColour,
        "borderRightColour" = borderRightColour,
        "borderBottomColour" = borderBottomColour,

        "halign" = halign,
        "valign" = valign,
        "indent" = indent,
        "textRotation" = textRotation,
        "numFmt" = numFmt,
        "fillFg" = fill$fillFg,
        "fillBg" = fill$fillBg,
        "wrapText" = wrapText,
        "locked" = locked,
        "hidden" = hidden,
        "xfId" = xfId
      )

      l[sapply(l, length) > 0]
    }
  )
)


mergeStyle <- function(oldStyle, newStyle) {

  ## This function is used to merge an existing cell style with a new style to create a stacked style.
  oldStyle <- oldStyle$copy()

  if (!is.null(newStyle$fontName)) {
    oldStyle$fontName <- newStyle$fontName
  }

  if (!is.null(newStyle$fontColour)) {
    oldStyle$fontColour <- newStyle$fontColour
  }

  if (!is.null(newStyle$fontSize)) {
    oldStyle$fontSize <- newStyle$fontSize
  }

  if (!is.null(newStyle$fontFamily)) {
    oldStyle$fontFamily <- newStyle$fontFamily
  }

  if (!is.null(newStyle$fontScheme)) {
    oldStyle$fontScheme <- newStyle$fontScheme
  }

  if (length(newStyle$fontDecoration) > 0) {
    if (length(oldStyle$fontDecoration) == 0) {
      oldStyle$fontDecoration <- newStyle$fontDecoration
    } else {
      oldStyle$fontDecoration <- c(oldStyle$fontDecoration, newStyle$fontDecoration)
    }
  }


  ## borders
  if (!is.null(newStyle$borderTop)) {
    oldStyle$borderTop <- newStyle$borderTop
  }

  if (!is.null(newStyle$borderLeft)) {
    oldStyle$borderLeft <- newStyle$borderLeft
  }

  if (!is.null(newStyle$borderRight)) {
    oldStyle$borderRight <- newStyle$borderRight
  }

  if (!is.null(newStyle$borderBottom)) {
    oldStyle$borderBottom <- newStyle$borderBottom
  }

  if (!is.null(newStyle$borderDiagonal)) {
    oldStyle$borderDiagonal <- newStyle$borderDiagonal
  }

  oldStyle$borderDiagonalUp <- newStyle$borderDiagonalUp
  oldStyle$borderDiagonalDown <- newStyle$borderDiagonalDown


  if (!is.null(newStyle$borderTopColour)) {
    oldStyle$borderTopColour <- newStyle$borderTopColour
  }

  if (!is.null(newStyle$borderLeftColour)) {
    oldStyle$borderLeftColour <- newStyle$borderLeftColour
  }

  if (!is.null(newStyle$borderRightColour)) {
    oldStyle$borderRightColour <- newStyle$borderRightColour
  }

  if (!is.null(newStyle$borderBottomColour)) {
    oldStyle$borderBottomColour <- newStyle$borderBottomColour
  }



  ## other
  if (!is.null(newStyle$halign)) {
    oldStyle$halign <- newStyle$halign
  }

  if (!is.null(newStyle$valign)) {
    oldStyle$valign <- newStyle$valign
  }

  if (!is.null(newStyle$indent)) {
    oldStyle$indent <- newStyle$indent
  }

  if (!is.null(newStyle$textRotation)) {
    oldStyle$textRotation <- newStyle$textRotation
  }

  if (!is.null(newStyle$numFmt)) {
    oldStyle$numFmt <- newStyle$numFmt
  }

  if (!is.null(newStyle$fill)) {
    oldStyle$fill <- newStyle$fill
  }

  if (!is.null(newStyle$wrapText)) {
    oldStyle$wrapText <- newStyle$wrapText
  }

  if (!is.null(newStyle$locked)) {
    oldStyle$locked <- newStyle$locked
  }

  if (!is.null(newStyle$hidden)) {
    oldStyle$hidden <- newStyle$hidden
  }

  if (!is.null(newStyle$xfId)) {
    oldStyle$xfId <- newStyle$xfId
  }

  return(oldStyle)
}


# wrapper -----------------------------------------------------------------

#' @name createStyle
#' @title Create a cell style
#' @description Create a new style to apply to worksheet cells
#' @author Alexander Walker
#' @seealso \code{\link{addStyle}}
#' @param fontName A name of a font. Note the font name is not validated. If fontName is NULL,
#' the workbook base font is used. (Defaults to Calibri)
#' @param fontColour Colour of text in cell.  A valid hex colour beginning with "#"
#' or one of colours(). If fontColour is NULL, the workbook base font colours is used.
#' (Defaults to black)
#' @param fontSize Font size. A numeric greater than 0.
#' If fontSize is NULL, the workbook base font size is used. (Defaults to 11)
#' @param numFmt Cell formatting
#' \itemize{
#'   \item{\bold{GENERAL}}
#'   \item{\bold{NUMBER}}
#'   \item{\bold{CURRENCY}}
#'   \item{\bold{ACCOUNTING}}
#'   \item{\bold{DATE}}
#'   \item{\bold{LONGDATE}}
#'   \item{\bold{TIME}}
#'   \item{\bold{PERCENTAGE}}
#'   \item{\bold{FRACTION}}
#'   \item{\bold{SCIENTIFIC}}
#'   \item{\bold{TEXT}}
#'   \item{\bold{COMMA}{  for comma separated thousands}}
#'   \item{For date/datetime styling a combination of d, m, y and punctuation marks}
#'   \item{For numeric rounding use "0.00" with the preferred number of decimal places}
#' }
#'
#' @param border Cell border. A vector of "top", "bottom", "left", "right" or a single string).
#' \itemize{
#'    \item{\bold{"top"}}{ Top border}
#'    \item{\bold{bottom}}{ Bottom border}
#'    \item{\bold{left}}{ Left border}
#'    \item{\bold{right}}{ Right border}
#'    \item{\bold{TopBottom} or \bold{c("top", "bottom")}}{ Top and bottom border}
#'    \item{\bold{LeftRight} or \bold{c("left", "right")}}{ Left and right border}
#'    \item{\bold{TopLeftRight} or \bold{c("top", "left", "right")}}{ Top, Left and right border}
#'    \item{\bold{TopBottomLeftRight} or \bold{c("top", "bottom", "left", "right")}}{ All borders}
#'   }
#'
#' @param borderColour Colour of cell border vector the same length as the number of sides specified in "border"
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"
#'
#' @param borderStyle Border line style vector the same length as the number of sides specified in "border"
#' \itemize{
#'    \item{\bold{none}}{ No Border}
#'    \item{\bold{thin}}{ thin border}
#'    \item{\bold{medium}}{ medium border}
#'    \item{\bold{dashed}}{ dashed border}
#'    \item{\bold{dotted}}{ dotted border}
#'    \item{\bold{thick}}{ thick border}
#'    \item{\bold{double}}{ double line border}
#'    \item{\bold{hair}}{ Hairline border}
#'    \item{\bold{mediumDashed}}{ medium weight dashed border}
#'    \item{\bold{dashDot}}{ dash-dot border}
#'    \item{\bold{mediumDashDot}}{ medium weight dash-dot border}
#'    \item{\bold{dashDotDot}}{ dash-dot-dot border}
#'    \item{\bold{mediumDashDotDot}}{ medium weight dash-dot-dot border}
#'    \item{\bold{slantDashDot}}{ slanted dash-dot border}
#'   }
#'
#' @param bgFill Cell background fill colour.
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#".
#' --  \bold{Use for conditional formatting styles only.}
#' @param fgFill Cell foreground fill colour.
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"
#'
#' @param halign
#' Horizontal alignment of cell contents
#' \itemize{
#'    \item{\bold{left}}{ Left horizontal align cell contents}
#'    \item{\bold{right}}{ Right horizontal align cell contents}
#'    \item{\bold{center}}{ Center horizontal align cell contents}
#'   }
#'
#' @param valign A name
#' Vertical alignment of cell contents
#' \itemize{
#'    \item{\bold{top}}{ Top vertical align cell contents}
#'    \item{\bold{center}}{ Center vertical align cell contents}
#'    \item{\bold{bottom}}{ Bottom vertical align cell contents}
#'   }
#'
#' @param textDecoration
#' Text styling.
#' \itemize{
#'    \item{\bold{bold}}{ Bold cell contents}
#'    \item{\bold{strikeout}}{ Strikeout cell contents}
#'    \item{\bold{italic}}{ Italicise cell contents}
#'    \item{\bold{underline}}{ Underline cell contents}
#'    \item{\bold{underline2}}{ Double underline cell contents}
#'   }
#'
#' @param wrapText Logical. If \code{TRUE} cell contents will wrap to fit in column.
#' @param textRotation Rotation of text in degrees. 255 for vertical text.
#' @param indent Horizontal indentation of cell contents.
#' @param hidden Whether the formula of the cell contents will be hidden (if worksheet protection is turned on)
#' @param locked Whether cell contents are locked (if worksheet protection is turned on)
#' @return A style object
#' @export
#' @examples
#' ## See package vignettes for further examples
#'
#' ## Modify default values of border colour and border line style
#' options("openxlsx.borderColour" = "#4F80BD")
#' options("openxlsx.borderStyle" = "thin")
#'
#' ## Size 18 Arial, Bold, left horz. aligned, fill colour #1A33CC, all borders,
#' style <- createStyle(
#'   fontSize = 18, fontName = "Arial",
#'   textDecoration = "bold", halign = "left", fgFill = "#1A33CC", border = "TopBottomLeftRight"
#' )
#'
#' ## Red, size 24, Bold, italic, underline, center aligned Font, bottom border
#' style <- createStyle(
#'   fontSize = 24, fontColour = rgb(1, 0, 0),
#'   textDecoration = c("bold", "italic", "underline"),
#'   halign = "center", valign = "center", border = "Bottom"
#' )
#'
#' # borderColour is recycled for each border or all colours can be supplied
#'
#' # colour is recycled 3 times for "Top", "Bottom" & "Right" sides.
#' createStyle(border = "TopBottomRight", borderColour = "red")
#'
#' # supply all colours
#' createStyle(border = "TopBottomLeft", borderColour = c("red", "yellow", "green"))
createStyle <- function(fontName = NULL,
  fontSize = NULL,
  fontColour = NULL,
  numFmt = "GENERAL",
  border = NULL,
  borderColour = getOption("openxlsx.borderColour", "black"),
  borderStyle = getOption("openxlsx.borderStyle", "thin"),
  bgFill = NULL, fgFill = NULL,
  halign = NULL, valign = NULL,
  textDecoration = NULL, wrapText = FALSE,
  textRotation = NULL,
  indent = NULL,
  locked = NULL, hidden = NULL) {

  ### Error checking
  od <- getOption("OutDec")
  options("OutDec" = ".")
  on.exit(expr = options("OutDec" = od), add = TRUE)

  ## if num fmt is made up of dd, mm, yy
  numFmt_original <- numFmt[[1]]
  numFmt <- tolower(numFmt_original)
  validNumFmt <- c("general", "number", "currency", "accounting", "date", "longdate", "time", "percentage", "scientific", "text", "3", "4", "comma")

  if (numFmt == "date") {
    numFmt <- getOption("openxlsx.dateFormat", getOption("openxlsx.dateformat", "date"))
  } else if (numFmt == "longdate") {
    numFmt <- getOption("openxlsx.datetimeFormat", getOption("openxlsx.datetimeformat", getOption("openxlsx.dateTimeFormat", "longdate")))
  } else if (!numFmt %in% validNumFmt) {
    numFmt <- replaceIllegalCharacters(numFmt_original)
  }




  numFmtMapping <- list(
    list("numFmtId" = 0), # GENERAL
    list("numFmtId" = 2), # NUMBER
    list("numFmtId" = 164, formatCode = "&quot;$&quot;#,##0.00"), ## CURRENCY
    list("numFmtId" = 44), # ACCOUNTING
    list("numFmtId" = 14), # DATE
    list("numFmtId" = 166, formatCode = "yyyy/mm/dd hh:mm:ss"), # LONGDATE
    list("numFmtId" = 167), # TIME
    list("numFmtId" = 10), # PERCENTAGE
    list("numFmtId" = 11), # SCIENTIFIC
    list("numFmtId" = 49), # TEXT

    list("numFmtId" = 3),
    list("numFmtId" = 4),
    list("numFmtId" = 3)
  )

  names(numFmtMapping) <- validNumFmt

  ## Validate border line style
  if (!is.null(borderStyle)) {
    borderStyle <- validateBorderStyle(borderStyle)
  }

  if (!is.null(halign)) {
    halign <- tolower(halign[[1]])
    if (!halign %in% c("left", "right", "center")) {
      stop("Invalid halign argument!")
    }
  }

  if (!is.null(valign)) {
    valign <- tolower(valign[[1]])
    if (!valign %in% c("top", "bottom", "center")) {
      stop("Invalid valign argument!")
    }
  }

  if (!is.logical(wrapText)) {
    stop("Invalid wrapText")
  }

  if (!is.null(indent)) {
    if (!is.numeric(indent) & !is.integer(indent)) {
      stop("indent must be numeric")
    }
  }

  textDecoration <- tolower(textDecoration)
  if (!is.null(textDecoration)) {
    if (!all(textDecoration %in% c("bold", "strikeout", "italic", "underline", "underline2", ""))) {
      stop("Invalid textDecoration!")
    }
  }

  borderColour <- validateColour(borderColour, "Invalid border colour!")

  if (!is.null(fontColour)) {
    fontColour <- validateColour(fontColour, "Invalid font colour!")
  }

  if (!is.null(fontSize)) {
    if (fontSize < 1) stop("Font size must be greater than 0!")
  }

  if (!is.null(locked)) {
    if (!is.logical(locked)) stop("Cell attribute locked must be TRUE or FALSE")
  }
  if (!is.null(hidden)) {
    if (!is.logical(hidden)) stop("Cell attribute hidden must be TRUE or FALSE")
  }





  ######################### error checking complete #############################
  style <- Style$new()

  if (!is.null(fontName)) {
    style$fontName <- list("val" = fontName)
  }

  if (!is.null(fontSize)) {
    style$fontSize <- list("val" = fontSize)
  }

  if (!is.null(fontColour)) {
    style$fontColour <- list("rgb" = fontColour)
  }

  style$fontDecoration <- toupper(textDecoration)

  ## background fill
  if (is.null(bgFill)) {
    bgFillList <- NULL
  } else {
    bgFill <- validateColour(bgFill, "Invalid bgFill colour")
    style$fill <- append(style$fill, list(fillBg = list("rgb" = bgFill)))
  }

  ## foreground fill
  if (is.null(fgFill)) {
    fgFillList <- NULL
  } else {
    fgFill <- validateColour(fgFill, "Invalid fgFill colour")
    style$fill <- append(style$fill, list(fillFg = list(rgb = fgFill)))
  }


  ## border
  if (!is.null(border)) {
    border <- toupper(border)
    border <- paste(border, collapse = "")

    ## find position of each side in string
    sides <- c("LEFT", "RIGHT", "TOP", "BOTTOM")
    pos <- sapply(sides, function(x) regexpr(x, border))
    pos <- pos[order(pos, decreasing = FALSE)]
    nSides <- sum(pos > 0)

    borderColour <- rep(borderColour, length.out = nSides)
    borderStyle <- rep(borderStyle, length.out = nSides)

    pos <- pos[pos > 0]

    if (length(pos) == 0) {
      stop("Unknown border argument")
    }

    names(borderColour) <- names(pos)
    names(borderStyle) <- names(pos)

    if ("LEFT" %in% names(pos)) {
      style$borderLeft <- borderStyle[["LEFT"]]
      style$borderLeftColour <- list("rgb" = borderColour[["LEFT"]])
    }

    if ("RIGHT" %in% names(pos)) {
      style$borderRight <- borderStyle[["RIGHT"]]
      style$borderRightColour <- list("rgb" = borderColour[["RIGHT"]])
    }

    if ("TOP" %in% names(pos)) {
      style$borderTop <- borderStyle[["TOP"]]
      style$borderTopColour <- list("rgb" = borderColour[["TOP"]])
    }

    if ("BOTTOM" %in% names(pos)) {
      style$borderBottom <- borderStyle[["BOTTOM"]]
      style$borderBottomColour <- list("rgb" = borderColour[["BOTTOM"]])
    }
  }

  ## other fields
  if (!is.null(halign)) {
    style$halign <- halign
  }

  if (!is.null(valign)) {
    style$valign <- valign
  }

  if (!is.null(indent)) {
    style$indent <- indent
  }

  if (wrapText) {
    style$wrapText <- TRUE
  }

  if (!is.null(textRotation)) {
    if (!is.numeric(textRotation)) {
      stop("textRotation must be numeric.")
    }

    if (textRotation < 0 & textRotation >= -90) {
      textRotation <- (textRotation * -1) + 90
    }

    style$textRotation <- round(textRotation[[1]], 0)
  }

  if (numFmt != "general") {
    if (numFmt %in% validNumFmt) {
      style$numFmt <- numFmtMapping[[numFmt[[1]]]]
    } else {
      style$numFmt <- list("numFmtId" = 165, formatCode = numFmt) ## Custom numFmt
    }
  }


  if (!is.null(locked)) {
    style$locked <- locked
  }

  if (!is.null(hidden)) {
    style$hidden <- hidden
  }

  return(style)
}

new_style <- function() {
  Style$new()
}
