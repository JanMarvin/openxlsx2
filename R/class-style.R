
# class -------------------------------------------------------------------

#' R6 class for a Workbook Style
#'
#' A style
#'
#' @export
wbStyle <- R6::R6Class(
  "wbStyle",
  public = list(
    # TODO consider what could just be a private field?

    # TODO can this just be a fontspecs() object?  with these values?

    #' @field fontId fontID
    fontId = NULL,
    #' @field fontName fontName
    fontName = NULL,
    #' @field fontColour fontColour
    fontColour = NULL,
    #' @field fontSize fontSize
    fontSize = NULL,
    #' @field fontFamily fontFamily
    fontFamily = NULL,
    #' @field fontScheme fontScheme
    fontScheme = NULL,
    #' @field fontDecoration fontDecoration
    fontDecoration = NULL,

    # TODO use border_specs() instead?

    #' @field borderTop borderTop
    borderTop = NULL,
    #' @field borderLeft borderLeft
    borderLeft = NULL,
    #' @field borderRight borderRight
    borderRight = NULL,
    #' @field borderBottom borderBottom
    borderBottom = NULL,
    #' @field borderTopColour borderTopColour
    borderTopColour = NULL,
    #' @field borderLeftColour borderLeftColour
    borderLeftColour = NULL,
    #' @field borderRightColour borderRightColour
    borderRightColour = NULL,
    #' @field borderBottomColour borderBottomColour
    borderBottomColour = NULL,
    #' @field borderDiagonal borderDiagonal
    borderDiagonal = NULL,
    #' @field borderDiagonalColour borderDiagonalColour
    borderDiagonalColour = NULL,
    #' @field borderDiagonalUp borderDiagonalUp
    borderDiagonalUp = NULL,
    #' @field borderDiagonalDown borderDiagonalDown
    borderDiagonalDown = NULL,


    #' @field halign Horizontal alignment
    halign = NULL,
    #' @field valign Vertical alignment
    valign = NULL,
    #' @field indent Indentation
    indent = NULL,
    #' @field textRotation The degree of text rotation
    textRotation = NULL,
    #' @field numFmt Number format id
    numFmt = NULL,
    #' @field fill fill
    fill = NULL,
    #' @field wrapText wrapText
    wrapText = NULL,
    #' @field locked locked
    locked = NULL,
    #' @field hidden hidden
    hidden = NULL,
    #' @field xfId xfId
    xfId = NULL,
    #' @field styleShow styleShow
    styleShow = NULL,

    #' @description
    #' Creates a new `wbStyle` object
    #' @param fontName fontName
    #' @param fontSize fontSize
    #' @param fontColour fontColour
    #' @param numFmt Number format id
    #' @param border Border
    #' @param borderColour border colour
    #' @param borderStyle border style
    #' @param bgFill Background fill
    #' @param fgFill Foreground fill
    #' @param halign Horizontal alignment
    #' @param valign Vertical alignment
    #' @param textDecoration text decorations
    #' @param wrapText If `TRUE` will wrap text
    #' @param textRotation The degree of text rotation
    #' @param indent indent
    #' @param locked locked
    #' @param hidden hidden
    #' @return a `wbStyle` object
    initialize = function(

      # TODO wb_font_spec() would be better
      fontName       = NULL,
      fontSize       = NULL,
      fontColour     = "none",
      numFmt         = "GENERAL",
      # numberFormat   = c("GENERAL", "NUMBER", "CURRENCY", "ACCOUNTING", "DATE", "LONGDATE", "TIME", "PERCENTAGE", "SCIENTIFIC", "TEXT", "3", "4", "COMMA"),
      # customFormat   = NULL, # this could be used to overwrite number format, and allow for easier match.arg()
      # if any border == none, sets to NULL
      # TODO wb_border_spec() would be better
      border         = c("none", "top", "bottom", "left", "right", "all"),
      borderColour   = getOption("openxlsx.borderColour", "black"),
      # should borderStyle default to "none"?
      borderStyle    = getOption("openxlsx.borderStyle", "thin"),
      bgFill         = "none",
      fgFill         = "none",
      halign         = c("left", "right", "center"),
      valign         = c("top", "center", "bottom"),
      # TODO should this be fontDecoration or textDecoration?
      textDecoration = c("none", "bold", "strikeout", "italic", "underline", "underline2"),
      wrapText       = FALSE,
      textRotation   = 0,
      indent         = NULL,
      locked         = FALSE,
      hidden         = FALSE
    ) {

      # asserts and validations ----
      assert_class(wrapText, "logical")
      assert_class(indent, c("numeric", "integer"), or_null = TRUE)
      assert_class(locked, "logical", or_null = TRUE)
      assert_class(hidden, "logical", or_null = TRUE)

      validFontColour <- validate_colour(fontColour, or_null = TRUE)
      bgFill          <- validate_colour(bgFill, or_null = TRUE)
      fgFill          <- validate_colour(fgFill, or_null = TRUE)


      # fonts ----
      self$fontName   <- value_list(fontName)
      self$fontSize   <- value_list(validate_font_size(fontSize))
      self$fontColour <- colour_list(validFontColour)

      self$halign         <- match.arg(halign)
      self$valign         <- match.arg(valign)
      text_style          <- match.arg(textDecoration, several.ok = TRUE)
      self$fontDecoration <- validate_text_style(text_style)
      self$numFmt         <- validate_number_format(numFmt)


      # fill ----
      self$fill <- c(
        if (!is.null(bgFill)) list(fillBg = list(rgb = bgFill)),
        if (!is.null(fgFill)) list(fillFg = list(rgb = fgFill))
      )

      # text rotation ----
      self$textRotation <- validate_text_rotation(textRotation)

      # borders ----
      private$set_borders(
        match.arg(border, several.ok = TRUE),
        colour = borderColour,
        style = borderStyle
      )

      # others ----
      self$indent   <- isTRUE(indent)
      self$locked   <- isTRUE(locked)
      self$hidden   <- isTRUE(hidden)
      self$wrapText <- isTRUE(wrapText)

      # auto set ----
      # when are these set?
      self$fontId               <- NULL
      self$fontFamily           <- NULL
      self$fontScheme           <- NULL
      self$borderDiagonal       <- NULL
      self$borderDiagonalColour <- NULL
      self$borderDiagonalUp     <- FALSE
      self$borderDiagonalDown   <- FALSE
      self$xfId                 <- NULL

      invisible(self)
    },

    # TODO as.list() to to_list() ?
    #' @description
    #' Convert the `wbStyle` object to a list
    #' @param include_null if `TRUE` will include values that contain `NULL`,
    #'   otherwise they are ommited from the list
    #' @return A `list` of fields
    as.list = function(include_null = FALSE) {
      # R6:::as.list.R6() exists...  would this be easier
      ls <- list(
        fontId         = self$fontId,
        fontName       = self$fontName,
        fontColour     = self$fontColour,
        fontSize       = self$fontSize,
        fontFamily     = self$fontFamily,
        fontScheme     = self$fontScheme,
        fontDecoration = self$fontDecoration,

        borderTop          = self$borderTop,
        borderLeft         = self$borderLeft,
        borderRight        = self$borderRight,
        borderBottom       = self$borderBottom,
        borderTopColour    = self$borderTopColour,
        borderLeftColour   = self$borderLeftColour,
        borderRightColour  = self$borderRightColour,
        borderBottomColour = self$borderBottomColour,

        halign       = self$halign,
        valign       = self$valign,
        indent       = self$indent,
        textRotation = self$textRotation,
        numFmt       = self$numFmt,
        fillFg       = self$fill$fillFg,
        fillBg       = self$fill$fillBg,
        wrapText     = self$wrapText,
        locked       = self$locked,
        hidden       = self$hidden,
        xfId         = self$xfId
      )

      if (include_null) {
        ls
      } else {
        ls[lengths(ls) > 0]
      }
    },

    #' @description
    #' Print the `wbStyle object`
    #' @param ... Not implemented
    #' @return the `wbStyle` object, invisibly; called for its side effects
    print = function(...) {
      .numFmtMapping <- c(
        GENERAL    =   0,
        NUMBER     =   2,
        CURRENCY   = 164,
        ACCOUNTING =  44,
        DATE       =  14,
        TIME       = 167,
        PERCENTAGE =  10,
        SCIENTIFIC =  11,
        TEXT       =  49
      )

      numFmtStr <- if (!is.null(self$numFmt)) {
        if (as.integer(self$numFmt$numFmtId) %in% .numFmtMapping) {
          names(.numFmtMapping)[.numFmtMapping == as.integer(self$numFmt$numFmtId)]
        } else {
          sprintf('"%s"', self$numFmt$formatCode)
        }
      } else {
        "GENERAL"
      }

      # if all are NULL
      borders <- c(
        sprintf("Top: %s",    self$borderTop),
        sprintf("Bottom: %s", self$borderBottom),
        sprintf("Left: %s",   self$borderLeft),
        sprintf("Right: %s",  self$borderRight)
      )

      borderColours <- c(self$borderTopColour, self$borderBottomColour, self$borderLeftColour, self$borderRightColour)
      borderColours <- gsub("^FF", "#", borderColours)

      # make as private
      self$styleShow  <- c(
        "A custom cell style. \n\n",
        # numFmt
        sprintf("Cell formatting: %s \n", numFmtStr),
        # Font name
        sprintf("Font name: %s \n", self$fontName[[1]]),
        # Font size
        sprintf("Font size: %s \n", self$fontSize[[1]]),
        # Font colour
        sprintf("Font colour: %s \n", gsub("^FF", "#", self$fontColour[[1]])),
        # Font decoration
        if (length(self$fontDecoration)) {
          sprintf("Font decoration: %s \n", paste(self$fontDecoration, collapse = ", "))
        },

        # Cell borders
        if (length(borders)) {
          c(
            sprintf("Cell borders: %s \n", paste(borders, collapse = ", ")),
            sprintf("Cell border colours: %s \n", paste(borderColours, collapse = ", "))
          )
        },

        # sprtinf("this %s", NULL) returns character()
        # Cell horizontal alignment
        sprintf("Cell horzizontal align: %s \n", self$halign),
        # Cell vertical alignment
        sprintf("Cell vertical align: %s \n", self$valign),
        # Cell indent
        sprintf("Cell indent: %s \n", self$indent),
        # Cell text rotation
        sprintf("Cell text rotation: %s \n", self$textRotation),
        # Cell fill colour
        if (length(self$fill$fillFg)) {
          sprintf(
            "Cell fill foreground: %s \n",
            paste(paste0(names(self$fill$fillFg), ": ", sub("^FF", "#", self$fill$fillFg)), collapse = ", ")
          )
        },
        # Cell background  fill
        if (length(self$fill$fillBg)) {
          sprintf(
            "Cell fill background: %s \n",
            paste(paste0(names(self$fill$fillBg), ": ", sub("^FF", "#", self$fill$fillBg)), collapse = ", ")
          )
        },
        # locked
        sprintf("Cell protection: %s \n", self$locked),
        # hidden
        sprintf("Cell formula hidden: %s \n", self$hidden),
        # wrapText
        sprintf("Cell wrap text: %s", self$wrapText),
        "\n\n"
      )

      # if (print) cat(self$styleShow)
      invisible(self)
    }
  ),

  private = list(
    set_borders = function(border, colour, style) {
      if (any(border == "none")) {
        # default will override
        self$borderLeft         <- NULL
        self$borderLeftColour   <- NULL
        self$borderRight        <- NULL
        self$borderRightColour  <- NULL
        self$borderTop          <- NULL
        self$borderTopColour    <- NULL
        self$borderBottom       <- NULL
        self$borderBottomColour <- NULL
        return(invisible(self))
      }

      if (identical(border, "all")) {
        pos <- c(left = 1L, right = 2L, top = 3L, bottom = 4L)
        n_sides <- 4L
      } else {

        ## find position of each side in string
        pos <- match(c("left", "right", "top", "bottom"), border, nomatch = 0L)
        names(pos) <- c("left", "right", "top", "bottom")
        pos <- sort(pos[pos > 0L])
        n_sides <- length(pos)

        if (!n_sides) {
          stop("border argument is not valid")
        }
      }

      colour <- rep.int(validate_colour(colour), n_sides)
      style  <- rep.int(validate_border_style(style), n_sides)

      names(colour) <- names(pos)
      names(style)  <- names(pos)

      # These are fine to be set to NULL
      self$borderLeft   <- na_to_null(style["left"])
      self$borderRight  <- na_to_null(style["right"])
      self$borderTop    <- na_to_null(style["top"])
      self$borderBottom <- na_to_null(style["bottom"])

      self$borderLeftColour   <- colour_list(colour["left"])
      self$borderRightColour  <- colour_list(colour["right"])
      self$borderTopColour    <- colour_list(colour["top"])
      self$borderBottomColour <- colour_list(colour["bottom"])

      invisible(self)
    }
  )
)


# wrappers ----------------------------------------------------------------

# TODO would this make sense as a method? wbStyle$merge(newStyle)?
mergeStyle <- function(oldStyle, newStyle) {
  assert_style(oldStyle)
  assert_style(newStyle)

  ## This function is used to merge an existing cell style with a new style to create a stacked style.
  res <- oldStyle$clone()

  for (i in merge_style_fields()) {
    res[[i]] <- newStyle[[i]] %||% oldStyle[[i]]
  }

  # separately
  res$fontDecoration <- c(oldStyle$fontDecoration, newStyle$fontDecoration)

  # do these need checks?
  res$borderDiagonalUp   <- newStyle$borderDiagonalUp
  res$borderDiagonalDown <- newStyle$borderDiagonalDown

  res
}

merge_style_fields <- function() {
  c(
    paste0("font"), c("Name", "Colour", "Size", "Family", "Scheme"),
    paste0("border", c("Top", "Left", "Right", "Bottom", "Diagonal")),
    paste0("border", c("Top", "Left", "Right", "Bottom"), "Colour"),
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
  )
}


# wrapper -----------------------------------------------------------------

wb_style <- function() {
  wbStyle$new()
}

#' @name createStyle
#' @title Create a cell style
#' @description Create a new style to apply to worksheet cells
#' @seealso [addStyle()]
#' @param fontName A name of a font. Note the font name is not validated. If fontName is NULL,
#' the workbook base font is used. (Defaults to Calibri)
#' @param fontColour Colour of text in cell.  A valid hex colour beginning with "#"
#' or one of colours(). If fontColour is NULL, the workbook base font colours is used.
#' (Defaults to black)
#' @param fontSize Font size. A numeric greater than 0.
#' If fontSize is NULL, the workbook base font size is used. (Defaults to 11)
#' @param numFmt Cell formatting
#' \itemize{
#'   \item{**GENERAL**}
#'   \item{**NUMBER**}
#'   \item{**CURRENCY**}
#'   \item{**ACCOUNTING**}
#'   \item{**DATE**}
#'   \item{**LONGDATE**}
#'   \item{**TIME**}
#'   \item{**PERCENTAGE**}
#'   \item{**FRACTION**}
#'   \item{**SCIENTIFIC**}
#'   \item{**TEXT**}
#'   \item{**COMMA**{  for comma separated thousands}}
#'   \item{For date/datetime styling a combination of d, m, y and punctuation marks}
#'   \item{For numeric rounding use "0.00" with the preferred number of decimal places}
#' }
#'
#' @param border Cell border. A vector of "top", "bottom", "left", "right" or a single string).
#' \itemize{
#'    \item{**none**}{ No border}
#'    \item{**"top"**}{ Top border}
#'    \item{**bottom**}{ Bottom border}
#'    \item{**left**}{ Left border}
#'    \item{**right**}{ Right border}
#'    \item{**all**}{ all borders}
#'   }
#'
#' @param borderColour Colour of cell border vector the same length as the number of sides specified in "border"
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"
#'
#' @param borderStyle Border line style vector the same length as the number of sides specified in "border"
#' \itemize{
#'    \item{**none**}{ No Border}
#'    \item{**thin**}{ thin border}
#'    \item{**medium**}{ medium border}
#'    \item{**dashed**}{ dashed border}
#'    \item{**dotted**}{ dotted border}
#'    \item{**thick**}{ thick border}
#'    \item{**double**}{ double line border}
#'    \item{**hair**}{ Hairline border}
#'    \item{**mediumDashed**}{ medium weight dashed border}
#'    \item{**dashDot**}{ dash-dot border}
#'    \item{**mediumDashDot**}{ medium weight dash-dot border}
#'    \item{**dashDotDot**}{ dash-dot-dot border}
#'    \item{**mediumDashDotDot**}{ medium weight dash-dot-dot border}
#'    \item{**slantDashDot**}{ slanted dash-dot border}
#'   }
#'
#' @param bgFill Cell background fill colour.
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#".
#' --  **Use for conditional formatting styles only.**
#' @param fgFill Cell foreground fill colour.
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"
#'
#' @param halign
#' Horizontal alignment of cell contents
#' \itemize{
#'    \item{**left**}{ Left horizontal align cell contents}
#'    \item{**right**}{ Right horizontal align cell contents}
#'    \item{**center**}{ Center horizontal align cell contents}
#'   }
#'
#' @param valign A name
#' Vertical alignment of cell contents
#' \itemize{
#'    \item{**top**}{ Top vertical align cell contents}
#'    \item{**center**}{ Center vertical align cell contents}
#'    \item{**bottom**}{ Bottom vertical align cell contents}
#'   }
#'
#' @param textDecoration
#' Text styling.
#' \itemize{
#'    \item{**none**}{ No cell contents}
#'    \item{**bold**}{ Bold cell contents}
#'    \item{**strikeout**}{ Strikeout cell contents}
#'    \item{**italic**}{ Italicise cell contents}
#'    \item{**underline**}{ Underline cell contents}
#'    \item{**underline2**}{ Double underline cell contents}
#'   }
#'
#' @param wrapText Logical. If `TRUE` cell contents will wrap to fit in column.
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
#' # options("openxlsx.borderColour" = "#4F80BD")
#' # options("openxlsx.borderStyle" = "thin")
#'
#' ## Size 18 Arial, Bold, left horz. aligned, fill colour #1A33CC, all borders,
#' style <- createStyle(
#'   fontSize = 18, fontName = "Arial",
#'   textDecoration = "bold", halign = "left", fgFill = "#1A33CC", border = "all"
#' )
#'
#' ## Red, size 24, Bold, italic, underline, center aligned Font, bottom border
#' style <- createStyle(
#'   fontSize = 24, fontColour = rgb(1, 0, 0),
#'   textDecoration = c("bold", "italic", "underline"),
#'   halign = "center", valign = "center", border = "bottom"
#' )
#'
#' # borderColour is recycled for each border or all colours can be supplied
#'
#' # colour is recycled 3 times for "Top", "Bottom" & "Right" sides.
#' createStyle(border = c("top", "bottom", "right"), borderColour = "red")
#'
#' # supply all colours
#' createStyle(border = c("top", "bottom", "left"), borderColour = c("red", "yellow", "green"))
createStyle <- function(
  fontName       = NULL,
  fontSize       = NULL,
  fontColour     = NULL,
  numFmt         = "GENERAL",
  border         = NULL,
  borderColour   = getOption("openxlsx.borderColour", "black"),
  borderStyle    = getOption("openxlsx.borderStyle", "thin"),
  bgFill         = NULL,
  fgFill         = NULL,
  halign         = c("left", "right", "center"),
  valign         = c("top", "center", "bottom"),
  # TODO should this be fontDecoration or textDecoration?
  textDecoration = c("none", "bold", "strikeout", "italic", "underline", "underline2"),
  wrapText       = FALSE,
  textRotation   = 0,
  indent         = NULL,
  locked         = FALSE,
  hidden         = FALSE
) {

  # TODO simplify options() setting
  od <- getOption("OutDec")
  options("OutDec" = ".")
  on.exit(expr = options("OutDec" = od), add = TRUE)


  wbStyle$new(
    fontName       = fontName,
    fontSize       = fontSize,
    fontColour     = fontColour,
    numFmt         = numFmt,
    border         = border,
    borderColour   = borderColour,
    borderStyle    = borderStyle,
    bgFill         = bgFill,
    fgFill         = fgFill,
    halign         = halign,
    valign         = valign,
    textDecoration = textDecoration,
    wrapText       = wrapText,
    textRotation   = textRotation,
    indent         = indent,
    locked         = locked,
    hidden         = hidden
  )
}


# helpers -----------------------------------------------------------------

number_formats <- function() {
  list(
    GENERAL     = list(numFmtId =   0L),
    NUMBER      = list(numFmtId =   2L),
    CURRENCY    = list(numFmtId = 164L, formatCode = "&quot;$&quot;#,##0.00"),
    ACCOUNTING  = list(numFmtId =  44L),
    DATE        = list(numFmtId =  14L),
    LONGDATE    = list(numFmtId = 166L, formatCode = "yyyy/mm/dd hh:mm:ss"),
    TIME        = list(numFmtId = 167L),
    PERCENTAGE  = list(numFmtId =  10L),
    SCIENTIFIC  = list(numFmtId =  11L),
    TEXT        = list(numFmtId =  49L),
    `3`         = list(numFmtId =   3L),
    `4`         = list(numFmtId =   4L),
    COMMA       = list(numFmtId =   3L)
  )
}

validate_text_style <- function(textDecoration = c("none", "bold", "strikeout", "italic", "underline", "underline2")) {
  x <- match.arg(textDecoration, several.ok = TRUE)
  x <- unique(x)

  # would consider warnings but then can't really use match.arg(); would need
  # something separate, maybe just the matching part without the errors

  if (any(x == "none")) {
    return(NULL)
  }

  if ("underline2" %in% x) {
    x <- x[x != "underline"]
  }

  x
}

validate_number_format <- function(x) {
  if (length(x) > 1) {
    stop("numFmt must be a single value")
  }

  x <- toupper(x)

  if (x %out% names(number_formats())) {
    x <- replaceIllegalCharacters(x)
  }

  x <- switch(
    x,
    # TODO make options more simple; use internal ox_options()
    date = getOption("openxlsx.dateFormat", getOption("openxlsx.dateformat", "date")),
    longdate = getOption("openxlsx.datetimeFormat", getOption("openxlsx.datetimeformat", getOption("openxlsx.dateTimeFormat", "longdate"))),
    x
  )

  number_formats()[[x]] %||% list(numFmtId = 165, formatCode = x)
}

validate_font_size <- function(fontSize) {
  if (is.null(fontSize)) {
    return(NULL)
  }

  assert_class(fontSize, c("numeric", "integer"))

  if (length(fontSize) > 1 || fontSize < 1) {
    stop("fontSize must be a single >= 1")
  }

  fontSize
}

validate_text_rotation <- function(x) {
  if (is.null(x)) {
    return(NULL)
  }

  if (!is.numeric(x) && length(x) == 1) {
    stop("textRotation must be single number")
  }

  if (x < 0 & x >= -90) {
    x <- (x * -1) + 90
  }

  round(x)
}

validate_border_style <- function(borderStyle = c("none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot")) {
  borderStyle <- match.arg(borderStyle, several.ok = TRUE)

  if (any(borderStyle == "none")) {
    return(NULL)
  }

  borderStyle
}


value_list <- function(x) {
  if (is.null(na_to_null(x))) {
    return(NULL)
  }

  list(val = x)
}

colour_list <- function(x) {
  if (is.null(na_to_null(x))) {
    return(NULL)
  }

  list(rgb = x)
}
