
# functions interacting with styles


#' clone sheets style
#'
#' @param wb workbook
#' @param from_sheet sheet we select the style from
#' @param to_sheet sheet we apply the style from
#' @export
cloneSheetStyle <- function(wb, from_sheet, to_sheet) {

  # check if sheets exist in wb
  id_org <- wb$validateSheet(from_sheet)
  id_new <- wb$validateSheet(to_sheet)

  org_style <- wb$worksheets[[id_org]]$sheet_data$cc
  wb_style <- wb$worksheets[[id_new]]$sheet_data$cc

  # remove all values
  org_style <- org_style[c("row_r", "c_r", "c_s")]

  merged_style <- merge(org_style, wb_style, all = TRUE)
  merged_style[is.na(merged_style)] <- "_openxlsx_NA_"

  # TODO order this as well?
  wb$worksheets[[id_new]]$sheet_data$cc <- merged_style

  # copy entire attributes from original sheet to new sheet
  org_rows <- wb$worksheets[[id_org]]$sheet_data$row_attr
  new_rows <- wb$worksheets[[id_new]]$sheet_data$row_attr

  merged_rows <- merge(org_rows[c("r")], new_rows, all = TRUE)
  merged_rows[is.na(merged_rows)] <- ""
  ordr <- ordered(order(as.integer(merged_rows$r)))
  merged_rows <- merged_rows[ordr,]

  wb$worksheets[[id_new]]$sheet_data$row_attr <- merged_rows

  wb$worksheets[[id_new]]$cols_attr <-
    wb$worksheets[[id_org]]$cols_attr

  wb$worksheets[[id_new]]$dimension <-
    wb$worksheets[[id_org]]$dimension

  wb$worksheets[[id_new]]$mergeCells <-
    wb$worksheets[[id_org]]$mergeCells

}


# internal function used in loadWorkbook
# @param x character string containing styles.xml
import_styles <- function(x) {

  sxml <- openxlsx2::read_xml(x)

  z <- NULL

  # Number Formats numFmtId
  # overrides for
  # + numFmtId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
  z$numFmts <- xml_node(sxml, "styleSheet", "numFmts", "numFmt")

  # Fonts fontId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1
  z$fonts <- xml_node(sxml, "styleSheet", "fonts", "font")

  # Fills fillId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fill?view=openxml-2.8.1
  z$fills <- xml_node(sxml, "styleSheet", "fills", "fill")

  # Borders borderId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-2.8.1
  z$borders <- xml_node(sxml, "styleSheet", "borders", "border")

  # Formatting Records: Cell Style Formats
  # links
  # + numFmtId
  # + fontId
  # + fillId
  # + borderId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-2.8.1
  z$cellStyleXfs <- xml_node(sxml, "styleSheet", "cellStyleXfs", "xf")

  # Cell Formats
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-2.8.1
  #
  # Position is 0 index s="value" used in worksheet <c ...>
  # links
  # + numFmtId
  # + fontId
  # + fillId
  # + borderId
  # + xfId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformats?view=openxml-2.8.1
  z$cellXfs <- xml_node(sxml, "styleSheet", "cellXfs", "xf")

  # Cell Styles (no clue)
  # links
  # + xfId
  # + builtinId
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle?view=openxml-2.8.1
  z$cellStyles <- xml_node(sxml, "styleSheet", "cellStyles", "cellStyle")

  # Differential Formats (No clue?)
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformat?view=openxml-2.8.1
  z$dxfs <- xml_node(sxml, "styleSheet", "dxfs", "dxf")

  # Table Styles Maybe position Id?
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyle?view=openxml-2.8.1
  z$tableStyles <- xml_node(sxml, "styleSheet", "tableStyles")

  # Colors
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colors?view=openxml-2.8.1
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-2.8.1
  z$colors <- xml_node(sxml, "styleSheet", "colors")

  # Future extensions No clue, some special styles
  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extensionlist?view=openxml-2.8.1
  z$extLst <- xml_node(sxml, "styleSheet", "extLst")

  z
}


#' get all styles on a sheet
#'
#' @param wb workbook
#' @param sheet worksheet
#'
#' @export
styles_on_sheet <- function(wb, sheet) {

  sheet_id <- wb$validateSheet(sheet)

  z <- unique(wb$worksheets[[sheet_id]]$sheet_data$cc$c_s)

  as.numeric(z)

}

# TODO guessing here
#' create fill
#'
#' @param diagonalDown x
#' @param diagonalUp x
#' @param outline x
#' @param bottom X
#' @param bottom_color X
#' @param diagonal X
#' @param diagonal_color X,
#' @param end x,
#' @param horizontal x
#' @param left x
#' @param left_color x
#' @param right x
#' @param right_color x
#' @param start x
#' @param top x
#' @param top_color x
#' @param vertical x
#'
#' @export
create_border <- function(
    diagonalDown = "",
    diagonalUp = "",
    outline = "",
    bottom = NULL,
    bottom_color = NULL,
    diagonal = NULL,
    diagonal_color = NULL,
    end = "",
    horizontal = "",
    left = NULL,
    left_color = NULL,
    right = NULL,
    right_color = NULL,
    start = "",
    top = NULL,
    top_color = NULL,
    vertical = ""
) {

  if (!is.null(left_color))     left_color     <- xml_node_create("color", xml_attributes = left_color)
  if (!is.null(right_color))    right_color    <- xml_node_create("color", xml_attributes = right_color)
  if (!is.null(top_color))      top_color      <- xml_node_create("color", xml_attributes = top_color)
  if (!is.null(bottom_color))   bottom_color   <- xml_node_create("color", xml_attributes = bottom_color)
  if (!is.null(diagonal_color)) diagonal_color <- xml_node_create("color", xml_attributes = diagonal_color)

  # excel dies on style=\"\"
  if (!is.null(left))     left     <- c(style = left)
  if (!is.null(right))    right    <- c(style = right)
  if (!is.null(top))      top      <- c(style = top)
  if (!is.null(bottom))   bottom   <- c(style = bottom)
  if (!is.null(diagonal)) diagonal <- c(style = diagonal)

  left     <- xml_node_create("left", xml_children = left_color, xml_attributes = left)
  right    <- xml_node_create("right", xml_children = right_color, xml_attributes = right)
  top      <- xml_node_create("top", xml_children = top_color, xml_attributes = top)
  bottom   <- xml_node_create("bottom", xml_children = bottom_color, xml_attributes = bottom)
  diagonal <- xml_node_create("diagonal", xml_children = diagonal_color, xml_attributes = diagonal)

  df_border <- data.frame(
    diagonalDown = diagonalDown,
    diagonalUp = diagonalUp,
    outline = outline,
    bottom = bottom,
    diagonal = diagonal,
    end = end,
    horizontal = horizontal,
    left = left,
    right = right,
    start = start,
    top = top,
    vertical = vertical,
    stringsAsFactors = FALSE
  )
  border <- write_border(df_border)

  return(border)
}


#' merge_borders
#' @param wb a workbook
#' @param new_borders <fill ...>
#' @export
merge_borders <- function(wb, new_borders) {

  # read old and new into dataframe
  if (length(wb$styles_mgr$styles$borders)) {
    old <- read_border(read_xml(wb$styles_mgr$styles$borders))
    new <- read_border(read_xml(new_borders))

    # get new rownames
    new_rownames <- seq(max(as.numeric(rownames(old)))+1,
                        length.out = NROW(new))
    row.names(new) <- as.character(new_rownames)

    # both have identical length, therefore can be rbind
    borders <- rbind(old, new)

    wb$styles_mgr$styles$borders <- write_border(borders)
  } else {
    wb$styles_mgr$styles$borders <- new_borders
    new_rownames <- seq_along(new_borders)
  }

  attr(wb, "new_borders") <- new_rownames
  # let the user now, which styles are new
  return(wb)
}

#' create number format
#' @param numFmtId an id
#' @param formatCode a format code
#' @export
create_numfmt <- function(numFmtId, formatCode) {

  df_numfmt <- data.frame(
    numFmtId = numFmtId,
    formatCode  = formatCode,
    stringsAsFactors = FALSE
  )
  numfmt <- write_numfmt(df_numfmt)

  return(numfmt)
}

# TODO should not be required to check for uniqueness of numFmts. This should
# have been done prior to calling this function.

#' merge_numFmts
#' @param wb a workbook
#' @param new_numfmts <numFmt ...>
#' @export
merge_numFmts <- function(wb, new_numfmts) {

  if (length(wb$styles_mgr$styles$numFmts)) {
    # read old and new into dataframe
    old <- read_numfmt(read_xml(wb$styles_mgr$styles$numFmts))
    new <- read_numfmt(read_xml(new_numfmts))

    # get new rownames
    new_rownames <- seq(max(as.numeric(rownames(old)))+1,
                        length.out = NROW(new))
    row.names(new) <- as.character(new_rownames)

    # both have identical length, therefore can be rbind
    numfmts <- rbind(old, new)

    wb$styles_mgr$styles$numFmts <- write_numfmt(numfmts)
  } else {
    wb$styles_mgr$styles$numFmts <- new_numfmts
    new_rownames <- seq_along(new_numfmts)
  }

  attr(wb, "new_numfmts") <- new_rownames
  # let the user now, which styles are new
  return(wb)
}

#' create number format
#' @param b bold
#' @param charset charset
#' @param color rgb color: default "FF000000"
#' @param condense condense
#' @param extend extend
#' @param family font family: default "2"
#' @param i italic
#' @param name font name: default "Calibri"
#' @param outline outline
#' @param scheme font scheme: default "minor"
#' @param shadow shadow
#' @param strike strike
#' @param sz font size: default "11",
#' @param u underline
#' @param vertAlign vertical alignment
#' @examples
#' \dontrun{
#' font <- create_font()
#' # openxml has the alpha value leading
#' hex8 <- unlist(xml_attr(read_xml(font), "font", "color"))
#' hex8 <- paste0("#", substr(hex8, 3, 8), substr(hex8, 1,2))
#'
#' # write test color
#' col <- crayon::make_style(col2rgb(hex8, alpha = TRUE))
#' cat(col("Test"))
#' }
#' @export
create_font <- function(
    b = "",
    charset = "",
    color = c(rgb = "FF000000"),
    condense ="",
    extend = "",
    family = "2",
    i = "",
    name = "Calibri",
    outline = "",
    scheme = "minor",
    shadow = "",
    strike = "",
    sz = "11",
    u = "",
    vertAlign = ""
) {

  if (b != "") {
    b <- xml_node_create("b", xml_attributes = c("val" = b))
  }

  if (charset != "") {
    charset <- xml_node_create("charset", xml_attributes = c("val" = charset))
  }

  if (color != "") {
    # alt xml_attributes(theme:)
    color <- xml_node_create("color", xml_attributes = color)
  }

  if (condense != "") {
    condense <- xml_node_create("condense", xml_attributes = c("val" = condense))
  }

  if (extend != "") {
    extend <- xml_node_create("extend", xml_attributes = c("val" = extend))
  }

  if (family != "") {
    family <- xml_node_create("family", xml_attributes = c("val" = family))
  }

  if (i != "") {
    i <- xml_node_create("i", xml_attributes = c("val" = i))
  }

  if (name != "") {
    name <- xml_node_create("name", xml_attributes = c("val" = name))
  }

  if (outline != "") {
    outline <- xml_node_create("outline", xml_attributes = c("val" = outline))
  }

  if (scheme != "") {
    scheme <- xml_node_create("scheme", xml_attributes = c("val" = scheme))
  }

  if (shadow != "") {
    shadow <- xml_node_create("shadow", xml_attributes = c("val" = shadow))
  }

  if (strike != "") {
    strike <- xml_node_create("strike", xml_attributes = c("val" = strike))
  }

  if (sz != "") {
    sz <- xml_node_create("sz", xml_attributes = c("val" = sz))
  }

  if (u != "") {
    u <- xml_node_create("u", xml_attributes = c("val" = u))
  }

  if (vertAlign != "") {
    vertAlign <- xml_node_create("vertAlign", xml_attributes = c("val" = vertAlign))
  }

  df_font <- data.frame(
    b = b,
    charset = charset,
    color = color,
    condense = condense,
    extend = extend,
    family = family,
    i = i,
    name = name,
    outline = outline,
    scheme = scheme,
    shadow = shadow,
    strike = strike,
    sz = sz,
    u = u,
    vertAlign = vertAlign,
    stringsAsFactors = FALSE
  )
  font <- write_font(df_font)

  return(font)
}

#' merge_fonts
#' @param wb a workbook
#' @param new_fonts <font ...>
#' @export
merge_fonts <- function(wb, new_fonts) {

  # read old and new into dataframe

  if (length(wb$styles_mgr$styles$fonts)) {
    old <- read_font(read_xml(wb$styles_mgr$styles$fonts))
    new <- read_font(read_xml(new_fonts))

    # get new rownames
    new_rownames <- seq(max(as.numeric(rownames(old)))+1,
                        length.out = NROW(new))
    row.names(new) <- as.character(new_rownames)

    # both have identical length, therefore can be rbind
    fonts <- rbind(old, new)

    wb$styles_mgr$styles$fonts <- write_font(fonts)
  } else {
    wb$styles_mgr$styles$fonts <- new_fonts
    new_rownames <- seq_along(new_fonts)
  }

  attr(wb, "new_fonts") <- new_rownames
  # let the user now, which styles are new
  return(wb)
}

#' create fill
#'
#' @param gradientFill complex fills
#' @param patternType various default is "none", but also "solid", or a color like "gray125"
#' @param bgColor hex8 color with alpha, red, green, blue only for patternFill
#' @param fgColor hex8 color with alpha, red, green, blue only for patternFill
#'
#' @export
create_fill <- function(
    gradientFill = "",
    patternType = "",
    bgColor = "",
    fgColor = ""
) {

  if (bgColor != ""){
    bgColor <- xml_node_create("bgColor", xml_attributes = bgColor)
  }

  if (fgColor != ""){
    fgColor <- xml_node_create("fgColor", xml_attributes = fgColor)
  }

  patternFill <- xml_node_create("patternFill",
                                 xml_children = c(bgColor, fgColor),
                                 xml_attributes = c(patternType = patternType))

  df_fill <- data.frame(
    gradientFill = gradientFill,
    patternFill = patternFill,
    stringsAsFactors = FALSE
  )
  fill <- write_fill(df_fill)

  return(fill)
}


#' merge_fills
#' @param wb a workbook
#' @param new_fills <fill ...>
#' @export
merge_fills <- function(wb, new_fills) {

  # read old and new into dataframe
  if (length(wb$styles_mgr$styles$fills)) {
    old <- read_fill(read_xml(wb$styles_mgr$styles$fills))
    new <- read_fill(read_xml(new_fills))

    # get new rownames
    new_rownames <- seq(max(as.numeric(rownames(old)))+1,
                        length.out = NROW(new))
    row.names(new) <- as.character(new_rownames)

    # both have identical length, therefore can be rbind
    fills <- rbind(old, new)

    wb$styles_mgr$styles$fills <- write_fill(fills)
  } else {
    wb$styles_mgr$styles$fills <- new_fills
    new_rownames <- seq_along(new_fills)
  }

  attr(wb, "new_fills") <- new_rownames
  # let the user now, which styles are new
  return(wb)
}

# TODO can be further generalized with additional xf attributes and children
#' create_cell_style
#' @param borderId dummy
#' @param fillId dummy
#' @param fontId dummy
#' @param numFmtId a numFmt ID for a builtin style
#' @param pivotButton dummy
#' @param quotePrefix dummy
#' @param xfId dummy
#' @param horizontal dummy
#' @param indent dummy
#' @param justifyLastLine dummy
#' @param readingOrder dummy
#' @param relativeIndent dummy
#' @param shrinkToFit dummy
#' @param textRotation dummy
#' @param vertical dummy
#' @param wrapText dummy
#' @param extLst dummy
#' @param hidden dummy
#' @param locked dummy
#' @param horizontal alignment can be "", "center", "right"
#' @param vertical alignment can be "", "center", "right"
#'
#' @details
#'  | "ID" | "numFmt"                    |
#'  |------|-----------------------------|
#'  | "0"  | "General"                   |
#'  | "1"  | "0"                         |
#'  | "2"  | "0.00"                      |
#'  | "3"  | "#,##0"                     |
#'  | "4"  | "#,##0.00"                  |
#'  | "9"  | "0%"                        |
#'  | "10" | "0.00%"                     |
#'  | "11" | "0.00E+00"                  |
#'  | "12" | "# ?/?"                     |
#'  | "13" | "# ??/??"                   |
#'  | "14" | "mm-dd-yy"                  |
#'  | "15" | "d-mmm-yy"                  |
#'  | "16" | "d-mmm"                     |
#'  | "17" | "mmm-yy"                    |
#'  | "18" | "h:mm AM/PM"                |
#'  | "19" | "h:mm:ss AM/PM"             |
#'  | "20" | "h:mm"                      |
#'  | "21" | "h:mm:ss"                   |
#'  | "22" | "m/d/yy h:mm"               |
#'  | "37" | "#,##0 ;(#,##0)"            |
#'  | "38" | "#,##0 ;\[Red\](#,##0)"     |
#'  | "39" | "#,##0.00;(#,##0.00)"       |
#'  | "40" | "#,##0.00;\[Red\](#,##0.00)"|
#'  | "45" | "mm:ss"                     |
#'  | "46" | "\[h\]:mm:ss"               |
#'  | "47" | "mmss.0"                    |
#'  | "48" | "##0.0E+0"                  |
#'  | "49" | "@"                         |
#' @export
create_cell_style <- function(
    borderId = "",
    fillId = "",
    fontId = "",
    numFmtId = "",
    pivotButton = "",
    quotePrefix = "",
    xfId = "",
    horizontal = "",
    indent = "",
    justifyLastLine = "",
    readingOrder = "",
    relativeIndent = "",
    shrinkToFit = "",
    textRotation = "",
    vertical = "",
    wrapText = "",
    extLst = "",
    hidden = "",
    locked = ""
) {
  n <- length(numFmtId)
  x <- as.character(numFmtId)

  applyAlignment <- ""
  if (any(horizontal != "") | any(textRotation != "") | any(vertical != "")) applyAlignment <- "1"

  applyBorder <- ""
  if (any(borderId != "")) applyBorder <- "1"

  applyFill <- ""
  if (any(fillId != "")) applyFill <- "1"

  applyFont <- ""
  if (any(fontId != "")) applyFont <- "1"

  applyNumberFormat <- ""
  if (any(x != "0")) applyNumberFormat <- "1"

  applyProtection <- ""
  if (any(hidden != "") | any(locked != "")) applyProtection <- "1"


  df_cellXfs <- data.frame(
    applyAlignment = rep(applyAlignment, n),
    applyBorder = rep(applyBorder, n),
    applyFill = rep(applyFill, n),
    applyFont = rep(applyFont, n),
    applyNumberFormat = rep(applyNumberFormat, n),
    applyProtection = rep(applyProtection, n),

    borderId = rep(borderId, n),
    fillId = rep(fillId, n),
    fontId = rep(fontId, n),
    numFmtId = numFmtId,
    pivotButton = rep(pivotButton, n),
    quotePrefix = rep(quotePrefix, n),
    xfId = rep(xfId, n),
    horizontal = rep(horizontal, n),
    indent = rep(indent, n),
    justifyLastLine = rep(justifyLastLine, n),
    readingOrder = rep(readingOrder, n),
    relativeIndent = rep(relativeIndent, n),
    shrinkToFit = rep(shrinkToFit, n),
    textRotation = rep(textRotation, n),
    vertical = rep(vertical, n),
    wrapText = rep(wrapText, n),
    extLst = rep(extLst, n),
    hidden = rep(hidden, n),
    locked = rep(locked, n),
    stringsAsFactors = FALSE
  )

  cellXfs <- write_xf(df_cellXfs)
  cellXfs
}

#' merge cellXfs styles from workbook and new
#' @param wb a workbook
#' @param new_cellxfs new character vector of <xf ...>
#' @export
merge_cellXfs <- function(wb, new_cellxfs) {

  # read old and new into dataframe
  old <- read_xf(read_xml(wb$styles_mgr$styles$cellXfs))
  new <- read_xf(read_xml(new_cellxfs))

  # get new rownames
  new_rownames <- seq(max(as.numeric(rownames(old)))+1,
                      length.out = NROW(new))
  row.names(new) <- as.character(new_rownames)

  # both have identical length, therefore can be rbind
  cellxfs <- rbind(old, new)

  wb$styles_mgr$styles$cellXfs <- write_xf(cellxfs)

  attr(wb, "wb_styles") <- new_rownames
  # let the user now, which styles are new
  return(wb)
}


#' helper get_cell_style
#' @param wb a workbook
#' @param sheet a worksheet
#' @param cell a cell
#' @export
get_cell_style <- function(wb, sheet, cell) {

  sheet <- wb$validateSheet(sheet)

  cell <- as.character(unlist(dims_to_dataframe(cell, fill = TRUE)))
  cc <- wb$worksheets[[sheet]]$sheet_data$cc

  cc$cell <- paste0(cc$c_r, cc$row_r)

  sel <- cc$cell %in% cell
  cc$c_s[sel]
}


#' helper set_cell_style
#' @param wb a workbook
#' @param sheet a worksheet
#' @param cell a cell
#' @param value a value to assign
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "test")
#'
#' mat <- matrix(rnorm(28*28, mean = 44444, sd = 555), ncol = 28)
#' writeData(wb, "test", mat, colNames = FALSE)
#'
#'
#' x <- c("0", "1", "2", "3", "4", "9", "10", "11", "12", "13", "14", "15", "16",
#'        "17", "18", "19", "20", "21", "22", "37", "38", "39", "40", "45", "46",
#'        "47", "48", "49")
#'
#' new_cellxfs <- create_cell_style(numFmtId = x, horizontal = "center")
#'
#' # combine original and new styles
#' wb <- merge_cellXfs(wb, new_cellxfs)
#' wb_styles <- attr(wb, "wb_styles")
#'
#' # new styles are 1:28, because s in a 0-index
#' for (i in wb_styles) {
#'   cell <- sprintf("%s1:%s28", int2col(i), int2col(i))
#'   set_cell_style(wb, "test", cell, as.character(i))
#' }
#'
#' \dontrun{
#' # look at the beauty you've created
#' wb_open(wb)
#' }
#' @export
set_cell_style <- function(wb, sheet, cell, value) {

  sheet <- wb$validateSheet(sheet)

  cell <- as.character(unlist(dims_to_dataframe(cell, fill = TRUE)))
  cc <- wb$worksheets[[sheet]]$sheet_data$cc

  cc$cell <- paste0(cc$c_r, cc$row_r)

  sel <- cc$cell %in% cell
  cc$c_s[sel] <- value
  cc$cell <- NULL

  wb$worksheets[[sheet]]$sheet_data$cc <- cc
}


# create_style <- function(wb, x) {
#
#
#
#   z <- list()
#
#   # Borders
#   z$borders <- character()
#
#   # Cell Styles
#   z$cellStyles <- character()
#
#   # Formatting Records
#   z$cellStyleXfs <- character()
#
#   # Cell Formats
#   z$cellXfs <- create_cell_style(x)
#
#   # Colors
#   z$colors <- character()
#
#   # Formats
#   z$dxfs <- character()
#
#   # Future extensions
#   z$extLst <- character()
#
#   # Fills
#   z$fills <- character()
#
#   # Fonts
#   z$fonts <- character()
#
#   # Number Formats
#   z$numFmts <- character()
#
#   # Table Styles
#   z$tableStyles <- character()
#
#   z
# }
