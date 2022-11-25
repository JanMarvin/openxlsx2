#' @name create_hyperlink
#' @title create Excel hyperlink string
#' @description Wrapper to create internal hyperlink string to pass to write_formula(). Either link to external urls or local files or straight to cells of local Excel sheets.
#' @param sheet Name of a worksheet
#' @param row integer row number for hyperlink to link to
#' @param col column number of letter for hyperlink to link to
#' @param text display text
#' @param file Excel file name to point to. If NULL hyperlink is internal.
#' @seealso [write_formula()]
#' @export create_hyperlink
#' @examples
#'
#' ## Writing internal hyperlinks
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet1")
#' wb$add_worksheet("Sheet2")
#' wb$add_worksheet("Sheet 3")
#' wb$add_data(sheet = 3, x = iris)
#'
#' ## External Hyperlink
#' x <- c("https://www.google.com", "https://www.google.com.au")
#' names(x) <- c("google", "google Aus")
#' class(x) <- "hyperlink"
#'
#' wb$add_data(sheet = 1, x = x, startCol = 10)
#'
#'
#' ## Internal Hyperlink - create hyperlink formula manually
#' write_formula(
#'   wb, "Sheet1",
#'   x = '=HYPERLINK(\"#Sheet2!B3\", "Text to Display - Link to Sheet2")',
#'   startCol = 3
#' )
#'
#' ## Internal - No text to display using create_hyperlink() function
#' write_formula(
#'   wb, "Sheet1",
#'   startRow = 1,
#'   x = create_hyperlink(sheet = "Sheet 3", row = 1, col = 2)
#' )
#'
#' ## Internal - Text to display
#' write_formula(
#'   wb, "Sheet1",
#'   startRow = 2,
#'   x = create_hyperlink(
#'     sheet = "Sheet 3", row = 1, col = 2,
#'     text = "Link to Sheet 3"
#'   )
#' )
#'
#' ## Link to file - No text to display
#' write_formula(
#'   wb, "Sheet1",
#'   startRow = 4,
#'   x = create_hyperlink(
#'     sheet = "testing", row = 3, col = 10,
#'     file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
#'   )
#' )
#'
#' ## Link to file - Text to display
#' write_formula(
#'   wb, "Sheet1",
#'   startRow = 3,
#'   x = create_hyperlink(
#'     sheet = "testing", row = 3, col = 10,
#'     file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"),
#'     text = "Link to File."
#'   )
#' )
#'
#' ## Link to external file - Text to display
#' write_formula(
#'   wb, "Sheet1",
#'   startRow = 10, startCol = 1,
#'   x = '=HYPERLINK("[C:/Users]", "Link to an external file")'
#' )
#'
#' ## Link to internal file
#' x = create_hyperlink(text = "test.png", file = "D:/somepath/somepicture.png")
#' write_formula(wb, "Sheet1", startRow = 11, startCol = 1, x = x)
create_hyperlink <- function(sheet, row = 1, col = 1, text = NULL, file = NULL) {
  if (missing(sheet)) {
    if (!missing(row) || !missing(col)) warning("Option for col and/or row found, but no sheet was provided.")

    if (is.null(text))
      str <- sprintf("=HYPERLINK(\"%s\")", file)

    if (is.null(file))
      str <- sprintf("=HYPERLINK(\"%s\")", text)

    if (!is.null(text) && !is.null(file))
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
  colour <- check_valid_colour(colour)

  if (isFALSE(colour)) {
    stop(errorMsg)
  }

  colour
}

check_valid_colour <- function(colour) {
  # Not proud of this.  Returns FALSE if not a vaild, otherwise  cleans up.
  # Probably not the best, but working within the functions we alreayd have.
  if (is.null(colour)) {
    colour <- "black"
  }

  validColours <- colours()

  if (any(colour %in% validColours)) {
    colour[colour %in% validColours] <- col2hex(colour[colour %in% validColours])
  }

  if (all(grepl("^#[A-Fa-f0-9]{6}$", colour))) {
    gsub("^#", "FF", toupper(colour))
  } else {
    FALSE
  }
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
    x <- replace_illegal_chars(x)
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


write_comment_xml <- function(comment_list, file_name) {
  authors <- unique(sapply(comment_list, "[[", "author"))
  xml <- '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
  xml <- c(xml, paste0("<authors>", paste(sprintf("<author>%s</author>", authors), collapse = ""), "</authors><commentList>"))

  for (i in seq_along(comment_list)) {
    authorInd <- which(authors == comment_list[[i]]$author) - 1L
    xml <- c(xml, sprintf('<comment ref="%s" authorId="%s" shapeId="0"><text>', comment_list[[i]]$ref, authorInd))

    ## Comment can have optional authors. Style and text is mandatory
    for (j in seq_along(comment_list[[i]]$comment)) {
      # write styles and comments
      if (is_xml(comment_list[[i]]$comment[[j]])) {
        comment <- comment_list[[i]]$comment[[j]]
      } else {
        comment <- sprintf('<t xml:space="preserve">%s</t>', comment_list[[i]]$comment[[j]])
      }

      xml <- c(xml, sprintf('<r>%s%s</r>',
        comment_list[[i]]$style[[j]],
        comment))
    }

    xml <- c(xml, "</text></comment>")
  }

  write_file(body = paste(xml, collapse = ""), tail = "</commentList></comments>", fl = file_name)

  NULL
}

pxml <- function(x) {
  ## TODO does this break anything? Why is unique called? lengths are off, if non unique values are found.
  # paste(unique(unlist(x)), collapse = "")
    paste(unlist(x), collapse = "")
}

#' split headerFooter xml into left, center, and right.
#' @param x xml string
#' @keywords internal
#' @noRd
amp_split <- function(x) {
  if (length(x) == 0) return(NULL)
  # create output string of width 3
  res <- vector("character", 3)
  # Identify the names found in the string: returns them as matrix: strip the &amp;
  nam <- gsub(pattern = "&amp;", "", unlist(stri_match_all_regex(x, "&amp;[LCR]")))
  # split the string and assign names to join
  z <- unlist(stri_split_regex(x, "&amp;[LCR]", omit_empty = TRUE))

  if (length(z) == 0) return(character(0))

  names(z) <- as.character(nam)
  res[c("L", "C", "R") %in% names(z)] <- z

  # return the string vector
  unname(res)
}

#' get headerFooter from xml into list with left, center, and right.
#' @param x xml string
#' @keywords internal
#' @noRd
getHeaderFooterNode <- function(x) {

  head_foot <- c("oddHeader", "oddFooter",
                  "evenHeader", "evenFooter",
                  "firstHeader", "firstFooter")

  headerFooter <- vector("list", length = length(head_foot))
  names(headerFooter) <- head_foot

  for (hf in head_foot) {
    headerFooter[[hf]] <- amp_split(xml_value(x, "headerFooter", hf))
  }

  headerFooter
}

#' generate headerFooter xml from left, center, and right characters
#' @param x xml string
#' @keywords internal
#' @noRd
genHeaderFooterNode <- function(x) {

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

  lcr <- function(x) {
    list(
      left   = ifelse(x[[1]] == "", "", sprintf("&amp;L%s", x[[1]])),
      center = ifelse(x[[2]] == "", "", sprintf("&amp;C%s", x[[2]])),
      right  = ifelse(x[[3]] == "", "", sprintf("&amp;R%s", x[[3]]))
    )
  }

  if (length(x$oddHeader)) {
    odd_header <- lcr(x$oddHeader)
    oddHeader <- paste0("<oddHeader>", odd_header[["left"]], odd_header[["center"]], odd_header[["right"]], "</oddHeader>", collapse = "")
  } else {
    oddHeader <- NULL
  }

  if (length(x$oddFooter)) {
    odd_footer <- lcr(x$oddFooter)
    oddFooter <- paste0("<oddFooter>", odd_footer[["left"]], odd_footer[["center"]], odd_footer[["right"]], "</oddFooter>", collapse = "")
  } else {
    oddFooter <- NULL
  }

  ## EVEN
  if (length(x$evenHeader)) {
    even_header <- lcr(x$evenHeader)
    evenHeader <- paste0("<evenHeader>", even_header[["left"]], even_header[["center"]], even_header[["right"]], "</evenHeader>", collapse = "")
  } else {
    evenHeader <- NULL
  }

  if (length(x$evenFooter)) {
    even_footer <- lcr(x$evenFooter)
    evenFooter <- paste0("<evenFooter>", even_footer[["left"]], even_footer[["center"]], even_footer[["right"]], "</evenFooter>", collapse = "")
  } else {
    evenFooter <- NULL
  }

  ## FIRST
  if (length(x$firstHeader)) {
    first_header <- lcr(x$firstHeader)
    firstHeader <- paste0("<firstHeader>", first_header[["left"]], first_header[["center"]], first_header[["right"]], "</firstHeader>", collapse = "")
  } else {
    firstHeader <- NULL
  }

  if (length(x$firstFooter)) {
    first_footer <- lcr(x$firstFooter)
    firstFooter <- paste0("<firstFooter>", first_footer[["left"]], first_footer[["center"]], first_footer[["right"]], "</firstFooter>", collapse = "")
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
  chars <- strsplit(substr(password, 1L, 15L), "")[[1]]

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

#' create sparklines used in `add_sparline()`
#' @details the colors are all predefined to be rgb. Maybe theme colors can be
#' used too.
#' @param sheet sheet
#' @param dims dims
#' @param sqref sqref
#' @param type type
#' @param negative negative
#' @param displayEmptyCellsAs displayEmptyCellsAs
#' @param markers markers add marker to line
#' @param high highlight highest value
#' @param low highlight lowest value
#' @param first highlight first value
#' @param last highlight last value
#' @param colorSeries colorSeries
#' @param colorNegative colorNegative
#' @param colorAxis colorAxis
#' @param colorMarkers colorMarkers
#' @param colorFirst colorFirst
#' @param colorLast colorLast
#' @param colorHigh colorHigh
#' @param colorLow colorLow
#' @examples
#' # create sparklineGroup
#' sparklines <- c(
#'   create_sparklines("Sheet 1", "A3:L3", "M3", type = "column", first = "1"),
#'   create_sparklines("Sheet 1", "A2:L2", "M2", markers = "1"),
#'   create_sparklines("Sheet 1", "A4:L4", "M4", type = "stacked", negative = "1")
#' )
#'
#' t1 <- AirPassengers
#' t2 <- do.call(cbind, split(t1, cycle(t1)))
#' dimnames(t2) <- dimnames(.preformat.ts(t1))
#'
#' wb <- wb_workbook()$
#'   add_worksheet("Sheet 1")$
#'   add_data(x = t2)$
#'   add_sparklines(sparklines = sparklines)
#'
#' @export
create_sparklines <- function(
    sheet = current_sheet(),
    dims,
    sqref,
    type = NULL,
    negative = NULL,
    displayEmptyCellsAs = "gap", # "span", "zero"
    markers = NULL,
    high = NULL,
    low = NULL,
    first = NULL,
    last = NULL,
    colorSeries = wb_colour(hex = "FF376092"),
    colorNegative = wb_colour(hex = "FFD00000"),
    colorAxis = wb_colour(hex = "FFD00000"),
    colorMarkers = wb_colour(hex = "FFD00000"),
    colorFirst = wb_colour(hex = "FFD00000"),
    colorLast = wb_colour(hex = "FFD00000"),
    colorHigh = wb_colour(hex = "FFD00000"),
    colorLow = wb_colour(hex = "FFD00000")
) {

  assert_class(dims, "character")
  assert_class(sqref, "character")

  ## FIXME validate_colour barks
  # colorSeries <- validate_colour(colorSeries)

  if (!is.null(type) && !type %in% c("stacked", "column"))
    stop("type must be NULL, stacked or column")

  if (!is.null(markers) && !is.null(type))
    stop("markers only work with stacked or column")


  sparklineGroup <- xml_node_create(
    "x14:sparklineGroup",
    xml_attributes = c(
      type = type,
      displayEmptyCellsAs = displayEmptyCellsAs,
      markers = markers,
      high = high,
      low = low,
      first = first,
      last = last,
      negative = negative,
      "xr2:uid" = sprintf("{6F57B887-24F1-C14A-942C-%s}", random_string(length = 12, pattern = "[A-Z0-9]"))
    ),
    xml_children = c(
      xml_node_create("x14:colorSeries", xml_attributes = colorSeries),
      xml_node_create("x14:colorNegative", xml_attributes = colorNegative),
      xml_node_create("x14:colorAxis", xml_attributes = colorAxis),
      xml_node_create("x14:colorMarkers", xml_attributes = colorMarkers),
      xml_node_create("x14:colorFirst", xml_attributes = colorFirst),
      xml_node_create("x14:colorLast", xml_attributes = colorLast),
      xml_node_create("x14:colorHigh", xml_attributes = colorHigh),
      xml_node_create("x14:colorLow", xml_attributes = colorLow),
      xml_node_create(
        "x14:sparklines", xml_children = c(
          xml_node_create(
            "x14:sparkline", xml_children = c(
              xml_node_create(
                "xm:f", xml_children = c(
                  paste0(shQuote(sheet, type = "sh"), "!", dims)
                )),
              xml_node_create(
                "xm:sqref", xml_children = c(
                  sqref
                ))
            ))
        )
      )
    )
  )

  sparklineGroup
}
