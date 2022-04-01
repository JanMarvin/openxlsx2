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
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet1")
#' wb$addWorksheet("Sheet2")
#' wb$addWorksheet("Sheet 3")
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
#'   x = '=HYPERLINK(\"#Sheet2!B3\", "Text to Display - Link to Sheet2")',
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
#' wb_save(wb, "internalHyperlinks.xlsx", overwrite = TRUE)
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

    ## Comment can have optional authors. Style and text is mandatory
    for (j in seq_along(comment_list[[i]]$comment)) {
      # write author to top of node. will be written in bold
      if (j == 1 & (comment_list[[i]]$author != ""))
        xml <- c(xml, sprintf('<r>%s<t xml:space="preserve">%s</t></r>',
          gsub("font>", "rPr>", create_font(b = "true")),
          paste0(comment_list[[i]]$author, ":\n")))

      # write styles and comments
      xml <- c(xml, sprintf('<r>%s<t xml:space="preserve">%s</t></r>',
        comment_list[[i]]$style[[j]],
        comment_list[[i]]$comment[[j]]))
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


get_named_regions_from_string <- function(dn) {
  dn <- cbind(
    openxlsx2:::rbindlist(xml_attr(dn, "definedName")),
    value =  xml_value(dn, "definedName")
  )

  if (!is.null(dn$value)) {
    dn_pos <- dn$value
    dn_pos <- gsub("[$']", "", dn_pos)

    has_bang <- grepl("!", dn_pos, fixed = TRUE)
    dn$sheets <- ifelse(has_bang, gsub("^(.*)!.*$", "\\1", dn_pos), "")
    dn$coords <- ifelse(has_bang, gsub("^.*!(.*)$", "\\1", dn_pos), "")
  }

  return(dn)
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
