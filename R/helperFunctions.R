#' @name makeHyperlinkString
#' @title create Excel hyperlink string
#' @description Wrapper to create internal hyperlink string to pass to writeFormula(). Either link to external urls or local files or straight to cells of local Excel sheets.
#' @param sheet Name of a worksheet
#' @param row integer row number for hyperlink to link to
#' @param col column number of letter for hyperlink to link to
#' @param text display text
#' @param file Excel file name to point to. If NULL hyperlink is internal.
#' @seealso [writeFormula()]
#' @export makeHyperlinkString
#' @examples
#'
#' ## Writing internal hyperlinks
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet1")
#' addWorksheet(wb, "Sheet2")
#' addWorksheet(wb, "Sheet 3")
#' writeData(wb, sheet = 3, x = iris)
#'
#' ## External Hyperlink
#' x <- c("https://www.google.com", "https://www.google.com.au")
#' names(x) <- c("google", "google Aus")
#' class(x) <- "hyperlink"
#'
#' writeData(wb, sheet = 1, x = x, startCol = 10)
#'
#'
#' ## Internal Hyperlink - create hyperlink formula manually
#' writeFormula(
#'   wb, "Sheet1",
#'   x = '=HYPERLINK(Sheet2!B3, "Text to Display - Link to Sheet2")',
#'   startCol = 3
#' )
#'
#' ## Internal - No text to display using makeHyperlinkString() function
#' writeFormula(
#'   wb, "Sheet1",
#'   startRow = 1,
#'   x = makeHyperlinkString(sheet = "Sheet 3", row = 1, col = 2)
#' )
#'
#' ## Internal - Text to display
#' writeFormula(
#'   wb, "Sheet1",
#'   startRow = 2,
#'   x = makeHyperlinkString(
#'     sheet = "Sheet 3", row = 1, col = 2,
#'     text = "Link to Sheet 3"
#'   )
#' )
#'
#' ## Link to file - No text to display
#' writeFormula(
#'   wb, "Sheet1",
#'   startRow = 4,
#'   x = makeHyperlinkString(
#'     sheet = "testing", row = 3, col = 10,
#'     file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
#'   )
#' )
#'
#' ## Link to file - Text to display
#' writeFormula(
#'   wb, "Sheet1",
#'   startRow = 3,
#'   x = makeHyperlinkString(
#'     sheet = "testing", row = 3, col = 10,
#'     file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"),
#'     text = "Link to File."
#'   )
#' )
#'
#' ## Link to external file - Text to display
#' writeFormula(
#'   wb, "Sheet1",
#'   startRow = 10, startCol = 1,
#'   x = '=HYPERLINK("[C:/Users]", "Link to an external file")'
#' )
#'
#' ## Link to internal file
#' x = makeHyperlinkString(text = "test.png", file = "D:/somepath/somepicture.png")
#' writeFormula(wb, "Sheet1", startRow = 11, startCol = 1, x = x)
#'
#' \dontrun{
#' saveWorkbook(wb, "internalHyperlinks.xlsx", overwrite = TRUE)
#' }
#'
makeHyperlinkString <- function(sheet, row = 1, col = 1, text = NULL, file = NULL) {
  # op <- get_set_options()
  # on.exit(options(op), add = TRUE)

  if (missing(sheet)) {
    if (!missing(row) || !missing(col)) warning("Option for col and/or row found, but no sheet was provided.")

    if (is.null(text))
      str <- sprintf("=HYPERLINK(\"%s\")", file)

    if (is.null(file))
      str <- sprintf("=HYPERLINK(\"%s\")", text)

    if (!is.null(text) & !is.null(file))
      str <- sprintf("=HYPERLINK(\"%s\", \"%s\")", file, text)
  } else {
    cell <- paste0(int2col(col), row)
    if (!is.null(file)) {
      dest <- sprintf('"[%s]%s!%s"', file, sheet, cell)
    } else {
      dest <- sprintf('"#\'%s\'!%s"', sheet, cell)
    }

    if (is.null(text)) {
      str <- sprintf('=HYPERLINK(%s)', dest)
    } else {
      str <- sprintf('=HYPERLINK(%s, \"%s\")', dest, text)
    }
  }

  return(str)
}


getRId <- function(x) reg_match0(x, '(?<= r:id=")[0-9A-Za-z]+')

getId <- function(x) reg_match0(x, '(?<= Id=")[0-9A-Za-z]+')



## creates style object based on column classes
## Used in writeData for styling when no borders and writeData table for all column-class based styling
classStyles <- function(wb, sheet, startRow, startCol, colNames, nRow, colClasses, stack = TRUE) {
  sheet <- wb$validateSheet(sheet)
  allColClasses <- unlist(colClasses, use.names = FALSE)
  rowInds <- (1 + startRow + colNames - 1L):(nRow + startRow + colNames - 1L)
  startCol <- startCol - 1L

  newStylesElements <- NULL
  names(colClasses) <- NULL

  if ("hyperlink" %in% allColClasses) {

    ## style hyperlinks
    inds <- wapply(colClasses, function(x) "hyperlink" %in% x)

    hyperlinkstyle <- createStyle(textDecoration = "underline")
    hyperlinkstyle$fontColour <- list("theme" = "10")
    styleElements <- list(
      "style" = hyperlinkstyle,
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }

  if ("date" %in% allColClasses) {

    ## style dates
    inds <- wapply(colClasses, function(x) "date" %in% x)

    styleElements <- list(
      "style" = createStyle(numFmt = "date"),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }

  if (any(c("posixlt", "posixct", "posixt") %in% allColClasses)) {

    ## style POSIX
    inds <- wapply(colClasses, function(x) any(c("posixct", "posixt", "posixlt") %in% x))

    styleElements <- list(
      "style" = createStyle(numFmt = "LONGDATE"),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }


  ## style currency as CURRENCY
  if ("currency" %in% allColClasses) {
    inds <- wapply(colClasses, function(x) "currency" %in% x)

    styleElements <- list(
      "style" = createStyle(numFmt = "CURRENCY"),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }

  ## style accounting as ACCOUNTING
  if ("accounting" %in% allColClasses) {
    inds <- wapply(colClasses, function(x) "accounting" %in% x)

    styleElements <- list(
      "style" = createStyle(numFmt = "ACCOUNTING"),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }

  ## style percentages
  if ("percentage" %in% allColClasses) {
    inds <- wapply(colClasses, function(x) "percentage" %in% x)

    styleElements <- list(
      "style" = createStyle(numFmt = "percentage"),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }

  ## style big mark
  if ("scientific" %in% allColClasses) {
    inds <- wapply(colClasses, function(x) "scientific" %in% x)

    styleElements <- list(
      "style" = createStyle(numFmt = "scientific"),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }

  ## style big mark
  if ("3" %in% allColClasses | "comma" %in% allColClasses) {
    inds <- wapply(colClasses, function(x) "3" %in% tolower(x) | "comma" %in% tolower(x))

    styleElements <- list(
      "style" = createStyle(numFmt = "3"),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }

  ## numeric sigfigs (Col must be numeric and numFmt options must only have 0s and \\.)
  if ("numeric" %in% allColClasses & !grepl("[^0\\.,#\\$\\* %]", getOption("openxlsx.numFmt", "GENERAL"))) {
    inds <- wapply(colClasses, function(x) "numeric" %in% tolower(x))

    styleElements <- list(
      "style" = createStyle(numFmt = getOption("openxlsx.numFmt", "0")),
      "sheet" = wb$sheet_names[sheet],
      "rows" = rep.int(rowInds, times = length(inds)),
      "cols" = rep(inds + startCol, each = length(rowInds))
    )

    newStylesElements <- c(newStylesElements, list(styleElements))
  }


  if (!is.null(newStylesElements)) {
    if (stack) {
      for (i in seq_along(newStylesElements)) {
        wb$addStyle(
          sheet = sheet,
          style = newStylesElements[[i]]$style,
          rows = newStylesElements[[i]]$rows,
          cols = newStylesElements[[i]]$cols, stack = TRUE
        )
      }
    } else {
      wb$styleObjects <- c(wb$styleObjects, newStylesElements)
    }
  }



  invisible(1)
}











#' @name validateColour
#' @description validate the colour input
#' @param colour colour
#' @param errorMsg Error message
#' @keywords internal
#' @noRd
validateColour <- function(colour, errorMsg = "Invalid colour!") {

  ## check if
  if (is.null(colour)) {
    colour <- "black"
  }

  validColours <- colours()

  if (any(colour %in% validColours)) {
    colour[colour %in% validColours] <- col2hex(colour[colour %in% validColours])
  }

  if (any(!grepl("^#[A-Fa-f0-9]{6}$", colour))) {
    stop(errorMsg, call. = FALSE)
  }

  colour <- gsub("^#", "FF", toupper(colour))

  return(colour)
}

#' @name col2hex
#' @description convert rgb to hex
#' @param creator my.col
#' @keywords internal
#' @noRd
col2hex <- function(my.col) {
  rgb(t(col2rgb(my.col)), maxColorValue = 255)
}


## header and footer replacements
headerFooterSub <- function(x) {
  if (!is.null(x)) {
    x <- replaceIllegalCharacters(x)
    x <- gsub("\\[Page\\]", "P", x)
    x <- gsub("\\[Pages\\]", "N", x)
    x <- gsub("\\[Date\\]", "D", x)
    x <- gsub("\\[Time\\]", "T", x)
    x <- gsub("\\[Path\\]", "Z", x)
    x <- gsub("\\[File\\]", "F", x)
    x <- gsub("\\[Tab\\]", "A", x)
  }

  return(x)
}


writeCommentXML <- function(comment_list, file_name) {
  authors <- unique(sapply(comment_list, "[[", "author"))
  xml <- '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
  xml <- c(xml, paste0("<authors>", paste(sprintf("<author>%s</author>", authors), collapse = ""), "</authors><commentList>"))

  for (i in seq_along(comment_list)) {
    authorInd <- which(authors == comment_list[[i]]$author) - 1L
    xml <- c(xml, sprintf('<comment ref="%s" authorId="%s" shapeId="0"><text>', comment_list[[i]]$ref, authorInd))

    if(length(comment_list[[i]]$style) != 0){ ## check that style information is present
      for (j in seq_along(comment_list[[i]]$comment)) {
        if (j == 1) # author
          xml <- c(xml, sprintf('<r>%s<t>%s</t></r>',
            comment_list[[i]]$style[[j]],
            comment_list[[i]]$comment[[j]]))
        if (j == 2) # comment
          xml <- c(xml, sprintf('<r>%s<t xml:space="preserve">%s</t></r>',
            comment_list[[i]]$style[[j]],
            comment_list[[i]]$comment[[j]]))
      }
    }else{ ## Case with no styling information.
      for (j in seq_along(comment_list[[i]]$comment)) {
        xml <- c(xml, sprintf('<t>%s</t>',
          comment_list[[i]]$comment[[j]]))
      }
    }

    xml <- c(xml, "</text></comment>")
  }

  write_file(body = paste(xml, collapse = ""), tail = "</commentList></comments>", fl = file_name)

  NULL
}


illegalchars <- c("&", '"', "'", "<", ">", "\a", "\b", "\v", "\f")
illegalcharsreplace <- c("&amp;", "&quot;", "&apos;", "&lt;", "&gt;", "", "", "", "")


replaceIllegalCharacters <- function(v) {
  vEnc <- Encoding(v)
  v <- as.character(v)

  flg <- vEnc != "UTF-8"
  if (any(flg)) {
    v[flg] <- stri_conv(v[flg], from = "", to = "UTF-8")
  }

  v <- stri_replace_all_fixed(v, illegalchars, illegalcharsreplace, vectorize_all = FALSE)

  return(v)
}


replaceXMLEntities <- function(v) {
  v <- gsub("&amp;", "&", v, fixed = TRUE)
  v <- gsub("&quot;", '"', v, fixed = TRUE)
  v <- gsub("&apos;", "'", v, fixed = TRUE)
  v <- gsub("&lt;", "<", v, fixed = TRUE)
  v <- gsub("&gt;", ">", v, fixed = TRUE)

  return(v)
}


pxml <- function(x) {
  ## TODO does this break anything? Why is unique called? lengths are off, if non unique values are found.
  # paste(unique(unlist(x)), collapse = "")
    paste(unlist(x), collapse = "")
}



buildFontList <- function(fonts) {

  sz     <- xml_attr(fonts, "font", "sz")
  colour <- xml_attr(fonts, "font", "color")
  name   <- xml_attr(fonts, "font", "name")
  family <- xml_attr(fonts, "font", "family")
  scheme <- xml_attr(fonts, "font", "scheme")

  italic <- lapply(fonts, function(x)xml_node(x, "font", "i"))
  bold <- lapply(fonts, function(x)xml_node(x, "font", "b"))
  underline <- lapply(fonts, function(x)xml_node(x, "font", "u"))

  ## Build font objects
  ft <- replicate(list(), n = length(fonts))
  for (i in seq_along(fonts)) {
    f <- NULL
    nms <- NULL
    if (length(unlist(sz[i]))) {
      f <- c(f, sz[i])
      nms <- c(nms, "sz")
    }

    if (length(unlist(colour[i]))) {
      f <- c(f, colour[i])
      nms <- c(nms, "color")
    }

    if (length(unlist(name[i]))) {
      f <- c(f, name[i])
      nms <- c(nms, "name")
    }

    if (length(unlist(family[i]))) {
      f <- c(f, family[i])
      nms <- c(nms, "family")
    }

    if (length(unlist(scheme[i]))) {
      f <- c(f, scheme[i])
      nms <- c(nms, "scheme")
    }

    if (length(italic[[i]])) {
      f <- c(f, "italic")
      nms <- c(nms, "italic")
    }

    if (length(bold[[i]])) {
      f <- c(f, "bold")
      nms <- c(nms, "bold")
    }

    if (length(underline[[i]])) {
      f <- c(f, "underline")
      nms <- c(nms, "underline")
    }

    f <- lapply(seq_along(f), function(i) unlist(f[i]))
    names(f) <- nms

    ft[[i]] <- f
  }

  ft
}



get_named_regions_from_string <- function(dn) {
  dn <- gsub("</definedNames>", "", dn, fixed = TRUE)
  dn <- gsub("</workbook>", "", dn, fixed = TRUE)

  dn <- unique(unlist(strsplit(dn, split = "</definedName>", fixed = TRUE)))
  dn <- grep("<definedName", dn, fixed = TRUE, value = TRUE)

  dn_names <- regmatches(dn, regexpr('(?<=name=")[^"]+', dn, perl = TRUE))

  dn_pos <- regmatches(dn, regexpr("(?<=>).*", dn, perl = TRUE))
  dn_pos <- gsub("[$']", "", dn_pos)

  has_bang <- grepl("!", dn_pos, fixed = TRUE)
  dn_sheets <- ifelse(has_bang, gsub("^(.*)!.*$", "\\1", dn_pos), "")
  dn_coords <- ifelse(has_bang, gsub("^.*!(.*)$", "\\1", dn_pos), "")

  attr(dn_names, "sheet") <- dn_sheets
  attr(dn_names, "position") <- dn_coords

  dn_names
}



nodeAttributes <- function(x) {
  x <- paste0("<", unlist(strsplit(x, split = "<")))
  x <- grep("<bgColor|<fgColor", x, value = TRUE)

  if (length(x) == 0) {
    return("")
  }

  attrs <- reg_match0(x, ' [a-zA-Z]+="[^"]*')
  tags <- reg_match0(x, "<[a-zA-Z]+ ")
  tags <- lapply(tags, gsub, pattern = "<| ", replacement = "")
  attrs <- lapply(attrs, gsub, pattern = '"', replacement = "")

  attrs <- lapply(attrs, strsplit, split = "=")
  for (i in seq_along(attrs)) {
    nms <- lapply(attrs[[i]], "[[", 1)
    vals <- lapply(attrs[[i]], "[[", 2)
    a <- unlist(vals)
    names(a) <- unlist(nms)
    attrs[[i]] <- a
  }

  names(attrs) <- unlist(tags)

  attrs
}


buildBorder <- function(x) {
  style <- list()
  if (grepl('diagonalup="1"', tolower(x), fixed = TRUE)) {
    style$borderDiagonalUp <- TRUE
  }

  if (grepl('diagonaldown="1"', tolower(x), fixed = TRUE)) {
    style$borderDiagonalDown <- TRUE
  }

  ## gets all borders that have children
  directions <- c("left", "right", "top", "bottom", "diagonal")
  x <- unapply(
    directions,
    function(z) {
      y <- xml_node(x, "border", z)
      sel <- length(xml_node(y, z, "color"))
      if (sel) y
    }
  )
  if (length(x) == 0) {
    return(NULL)
  }

  sides <- c("TOP", "BOTTOM", "LEFT", "RIGHT", "DIAGONAL")
  sideBorder <- character(length = length(x))
  for (i in seq_along(x)) {
    tmp <- sides[sapply(sides, function(s) grepl(s, x[[i]], ignore.case = TRUE))]
    if (length(tmp) > 1) tmp <- tmp[[1]]
    if (length(tmp) == 1) {
      sideBorder[[i]] <- tmp
    }
  }

  sideBorder <- sideBorder[sideBorder != ""]
  x <- x[sideBorder != ""]
  if (length(sideBorder) == 0) {
    return(NULL)
  }


  ## style
  weight <- gsub('style=|"', "", regmatches(x, regexpr('style="[a-z]+"', x, perl = TRUE, ignore.case = TRUE)))


  ## Colours
  cols <- replicate(n = length(sideBorder), list(rgb = "FF000000"))
  colNodes <- sapply(x, function(xi) xml_node(xi, "border", "*", "color"))


  if (length(colNodes)) {
    attrs <- regmatches(colNodes, regexpr('(theme|indexed|rgb|auto)=".+"', colNodes))
  } else {
    attrs <- NULL
  }

  if (length(attrs) != length(x)) {
    return(
      list(
        "borders" = paste(sideBorder, collapse = ""),
        "colour" = cols
      )
    )
  }

  attrs <- strsplit(attrs, split = "=")
  cols <- sapply(
    attrs,
    function(attr) {
      if (length(attr) == 2) {
        y <- list(gsub('"', "", attr[2]))
        names(y) <- gsub(" ", "", attr[[1]])
      } else {
        tmp <- paste(attr[-1], collapse = "=")
        y <- gsub('^"|"$', "", tmp)
        names(y) <- gsub(" ", "", attr[[1]])
      }
      y
    }
  )

  ## sideBorder & cols
  if ("LEFT" %in% sideBorder) {
    style$borderLeft <- weight[which(sideBorder == "LEFT")]
    style$borderLeftColour <- cols[which(sideBorder == "LEFT")]
  }

  if ("RIGHT" %in% sideBorder) {
    style$borderRight <- weight[which(sideBorder == "RIGHT")]
    style$borderRightColour <- cols[which(sideBorder == "RIGHT")]
  }

  if ("TOP" %in% sideBorder) {
    style$borderTop <- weight[which(sideBorder == "TOP")]
    style$borderTopColour <- cols[which(sideBorder == "TOP")]
  }

  if ("BOTTOM" %in% sideBorder) {
    style$borderBottom <- weight[which(sideBorder == "BOTTOM")]
    style$borderBottomColour <- cols[which(sideBorder == "BOTTOM")]
  }

  if ("DIAGONAL" %in% sideBorder) {
    style$borderDiagonal <- weight[which(sideBorder == "DIAGONAL")]
    style$borderDiagonalColour <- cols[which(sideBorder == "DIAGONAL")]
  }

  return(style)
}




genHeaderFooterNode <- function(x) {
  # TODO is x some class of something?

  # <headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="0">
  #   <oddHeader>&amp;Lfirst L&amp;CfC&amp;RfR</oddHeader>
  #   <oddFooter>&amp;LfFootL&amp;CfFootC&amp;RfFootR</oddFooter>
  #   <evenHeader>&amp;LTIS&amp;CIS&amp;REVEN H</evenHeader>
  #   <evenFooter>&amp;LEVEN L F&amp;CEVEN C F&amp;REVEN RIGHT F</evenFooter>
  #   <firstHeader>&amp;L&amp;P&amp;Cfirst C&amp;Rfirst R</firstHeader>
  #   <firstFooter>&amp;Lfirst L Foot&amp;Cfirst C Foot&amp;Rfirst R Foot</firstFooter>
  #   </headerFooter>

  ## ODD

  # TODO clean up length(x)
  # TODO clean up to x <- if (cond) value; default here is NULL

  # return nothing if there is no length
  if (!length(x)) return(NULL)

  if (length(x$oddHeader)) {
    oddHeader <- paste0("<oddHeader>", sprintf("&amp;L%s", x$oddHeader[[1]]), sprintf("&amp;C%s", x$oddHeader[[2]]), sprintf("&amp;R%s", x$oddHeader[[3]]), "</oddHeader>", collapse = "")
  } else {
    oddHeader <- NULL
  }

  if (length(x$oddFooter)) {
    oddFooter <- paste0("<oddFooter>", sprintf("&amp;L%s", x$oddFooter[[1]]), sprintf("&amp;C%s", x$oddFooter[[2]]), sprintf("&amp;R%s", x$oddFooter[[3]]), "</oddFooter>", collapse = "")
  } else {
    oddFooter <- NULL
  }

  ## EVEN
  if (length(x$evenHeader)) {
    evenHeader <- paste0("<evenHeader>", sprintf("&amp;L%s", x$evenHeader[[1]]), sprintf("&amp;C%s", x$evenHeader[[2]]), sprintf("&amp;R%s", x$evenHeader[[3]]), "</evenHeader>", collapse = "")
  } else {
    evenHeader <- NULL
  }

  if (length(x$evenFooter)) {
    evenFooter <- paste0("<evenFooter>", sprintf("&amp;L%s", x$evenFooter[[1]]), sprintf("&amp;C%s", x$evenFooter[[2]]), sprintf("&amp;R%s", x$evenFooter[[3]]), "</evenFooter>", collapse = "")
  } else {
    evenFooter <- NULL
  }

  ## FIRST
  if (length(x$firstHeader)) {
    firstHeader <- paste0("<firstHeader>", sprintf("&amp;L%s", x$firstHeader[[1]]), sprintf("&amp;C%s", x$firstHeader[[2]]), sprintf("&amp;R%s", x$firstHeader[[3]]), "</firstHeader>", collapse = "")
  } else {
    firstHeader <- NULL
  }

  if (length(x$firstFooter)) {
    firstFooter <- paste0("<firstFooter>", sprintf("&amp;L%s", x$firstFooter[[1]]), sprintf("&amp;C%s", x$firstFooter[[2]]), sprintf("&amp;R%s", x$firstFooter[[3]]), "</firstFooter>", collapse = "")
  } else {
    firstFooter <- NULL
  }


  headTag <- sprintf(
    '<headerFooter differentOddEven="%s" differentFirst="%s" scaleWithDoc="0" alignWithMargins="0">',
    as.integer(!(is.null(evenHeader) & is.null(evenFooter))),
    as.integer(!(is.null(firstHeader) & is.null(firstFooter)))
  )

  paste0(headTag, oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter, "</headerFooter>")
}


buildFillList <- function(fills) {
  fillAttrs <- rep(list(list()), length(fills))

  ## patternFill
  inds <- grepl("patternFill", fills)
  fillAttrs[inds] <- lapply(fills[inds], nodeAttributes)

  ## gradientFill
  inds <- grepl("gradientFill", fills)
  fillAttrs[inds] <- fills[inds]

  return(fillAttrs)
}


getDefinedNamesSheet <- function(x) {
  # belongTo <- unlist(lapply(strsplit(x, split = ">|<"), "[[", 3))
  # quoted <- grepl("^'", belongTo)

  # belongTo[quoted] <- regmatches(belongTo[quoted], regexpr("(?<=').*(?='!)", belongTo[quoted], perl = TRUE))
  # belongTo[!quoted] <- gsub("!\\$[A-Z0-9].*", "", belongTo[!quoted])
  # belongTo[!quoted] <- gsub("!#REF!.*", "", belongTo[!quoted])

  # return(belongTo)
}


clean_names <- function(x, schar) {
  x <- gsub("^[[:space:]]+|[[:space:]]+$", "", x)
  x <- gsub("[[:space:]]+", schar, x)
  return(x)
}



mergeCell2mapping <- function(x) {
  refs <- regmatches(x, regexpr("(?<=ref=\")[A-Z0-9:]+", x, perl = TRUE))
  refs <- strsplit(refs, split = ":")
  rows <- lapply(refs, function(r) {
    r <- as.integer(gsub(pattern = "[A-Z]", replacement = "", r, perl = TRUE))
    seq(from = r[1], to = r[2], by = 1)
  })

  cols <- lapply(refs, function(r) {
    r <- col2int(r)
    seq(from = r[1], to = r[2], by = 1)
  })

  ## for each we grid.expand
  refs <- do.call("rbind", lapply(seq_along(rows), function(i) {
    tmp <- expand.grid("cols" = cols[[i]], "rows" = rows[[i]])
    tmp$ref <- paste0(int2col(tmp$cols), tmp$rows)
    tmp$anchor_cell <- tmp$ref[1]
    return(tmp[, c("anchor_cell", "ref", "rows")])
  }))


  refs <- refs[refs$anchor_cell != refs$ref, ]

  return(refs)
}



getFile <- function(xlsxFile) {

  ## Is this a file or URL (code taken from read.table())
  on.exit(try(close(fl), silent = TRUE), add = TRUE)
  fl <- file(description = xlsxFile)

  ## If URL download
  if (inherits(fl, "url")) {
    tmpFile <- tempfile(fileext = ".xlsx")
    download.file(url = xlsxFile, destfile = tmpFile, cacheOK = FALSE, mode = "wb", quiet = TRUE)
    xlsxFile <- tmpFile
  }

  return(xlsxFile)
}

# Rotate the 15-bit integer by n bits to the
hashPassword <- function(password) {
  # password limited to 15 characters
  # TODO add warning about password length
  # TODO use substr(password, 1, 15) instead
  chars <- head(strsplit(password, "")[[1]], 15)
  # See OpenOffice's documentation of the Excel format: http://www.openoffice.org/sc/excelfileformat.pdf
  # Start from the last character and for each character
  # - XOR hash with the ASCII character code
  # - rotate hash (16 bits) one bit to the left
  # Finally, XOR hash with 0xCE4B and XOR with password length
  # Output as hex (uppercase)
  rotate16bit <- function(hash, n = 1) {
    bitwOr(bitwAnd(bitwShiftR(hash, 15 - n), 0x01), bitwAnd(bitwShiftL(hash, n), 0x7fff))
  }
  hash <- Reduce(function(char, h) {
    h <- bitwXor(h, as.integer(charToRaw(char)))
    rotate16bit(h, 1)
  }, chars, 0, right = TRUE)
  hash <- bitwXor(bitwXor(hash, length(chars)), 0xCE4B)
  format(as.hexmode(hash), upper.case = TRUE)
}
