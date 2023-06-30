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
#'     file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#'   )
#' )
#'
#' ## Link to file - Text to display
#' write_formula(
#'   wb, "Sheet1",
#'   startRow = 3,
#'   x = create_hyperlink(
#'     sheet = "testing", row = 3, col = 10,
#'     file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"),
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


#' @name validateColor
#' @description validate the color input
#' @param color color
#' @param errorMsg Error message
#' @keywords internal
#' @noRd
validateColor <- function(color, errorMsg = "Invalid color!") {
  color <- check_valid_color(color)

  if (isFALSE(color)) {
    stop(errorMsg)
  }

  color
}

check_valid_color <- function(color) {
  # Not proud of this.  Returns FALSE if not a vaild, otherwise  cleans up.
  # Probably not the best, but working within the functions we alreayd have.
  if (is.null(color)) {
    color <- "black"
  }

  validColors <- colors()

  if (any(color %in% validColors)) {
    color[color %in% validColors] <- col2hex(color[color %in% validColors])
  }

  if (all(grepl("^#[A-Fa-f0-9]{6}$", color))) {
    gsub("^#", "FF", toupper(color))
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

      is_fmt_txt <- FALSE
      if (is_xml(comment))
      is_fmt_txt <- all(xml_node_name(comment) == "r")

      if (is_fmt_txt) {
        xml <- c(xml, comment)
      } else {
        xml <- c(xml, sprintf('<r>%s%s</r>',
                              comment_list[[i]]$style[[j]],
                              comment))
      }
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
    colorSeries = wb_color(hex = "FF376092"),
    colorNegative = wb_color(hex = "FFD00000"),
    colorAxis = wb_color(hex = "FFD00000"),
    colorMarkers = wb_color(hex = "FFD00000"),
    colorFirst = wb_color(hex = "FFD00000"),
    colorLast = wb_color(hex = "FFD00000"),
    colorHigh = wb_color(hex = "FFD00000"),
    colorLow = wb_color(hex = "FFD00000")
) {

  assert_class(dims, "character")
  assert_class(sqref, "character")

  ## FIXME validate_color barks
  # colorSeries <- validate_color(colorSeries)

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

### beg pivot table helpers

cacheFields <- function(wbdata, filter, rows, cols, data) {
  sapply(
    names(wbdata),
    function(x) {

      dat <- wbdata[[x]]

      vars <- c(filter, rows, cols, data)

      is_vars <- x %in% vars
      is_data <- x %in% data &&
        # there is an exception to every rule in pivot tables ...
        # if a data variable is col or row we need items
        !x %in% cols && !x %in% rows && !x %in% filter
      is_char <- is.character(dat)
      is_date <- inherits(dat, "Date") || inherits(dat, "POSIXct")
      if (is_date) {
        dat <- format(dat, format = "%Y-%m-%dT%H:%M:%S")
      }
      unis <- stringi::stri_unique(dat)

      if (is_vars && !is_data) {
        sharedItem <- vapply(
          unis,
          function(uni) {
            if (is.na(uni)) {
              xml_node_create("m")
            } else {

              if (is_char) {
                xml_node_create("s", xml_attributes = c(v = uni), escapes = TRUE)
              } else if (is_date) {
                xml_node_create("d", xml_attributes = c(v = uni))
              } else {
                xml_node_create("n", xml_attributes = c(v = uni))
              }
            }

          },
          FUN.VALUE = ""
        )
      } else {
        sharedItem <- NULL
      }

      if (any(is.na(dat))) {
        containsBlank <- "1"
        containsSemiMixedTypes <- NULL
      } else {
        containsBlank <- NULL
        # indicates that cell has no blank values
        containsSemiMixedTypes <- "0"
      }

      if (length(sharedItem)) {
        count <- as.character(length(sharedItem))
      } else {
        count <- NULL
      }

      if (is_char) {
        attr <- c(
          containsBlank = containsBlank,
          count = count
        )
      } else if (is_date) {
        attr <- c(
          containsNonDate = "0",
          containsDate = "1",
          # enabled, check side effect
          containsString = "0",
          containsSemiMixedTypes = containsSemiMixedTypes,
          containsBlank = containsBlank,
          # containsMixedTypes = "1",
          minDate = format(min(dat, na.rm = TRUE), format = "%Y-%m-%dT%H:%M:%S"),
          maxDate = format(max(dat, na.rm = TRUE), format = "%Y-%m-%dT%H:%M:%S"),
          count = count
        )

      } else {

        # numeric or integer
        is_int <- all(dat[!is.na(dat)] %% 1 == 0)
        if (is_int) {
          containsInteger <- as_binary(is_int)
        } else {
          containsInteger <- NULL
        }

        attr <- c(
          containsSemiMixedTypes = containsSemiMixedTypes,
          containsString = "0",
          containsBlank = containsBlank,
          containsNumber = "1",
          containsInteger = containsInteger, # double or int?
          minValue = as.character(min(dat, na.rm = TRUE)),
          maxValue = as.character(max(dat, na.rm = TRUE)),
          count = count
        )
      }

      sharedItems <- xml_node_create(
        "sharedItems",
        xml_attributes = attr,
        xml_children = sharedItem
      )

      xml_node_create(
        "cacheField",
        xml_attributes = c(name = x, numFmtId = "0"),
        xml_children = sharedItems
      )
    },
    USE.NAMES = FALSE
  )
}

pivot_def_xml <- function(wbdata, filter, rows, cols, data) {

  ref   <- dataframe_to_dims(attr(wbdata, "dims"))
  sheet <- attr(wbdata, "sheet")
  count <- ncol(wbdata)

  paste0(
    sprintf('<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="xr" r:id="rId1" refreshedBy="openxlsx2" invalid="1" refreshOnLoad="1" refreshedDate="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="%s">', nrow(wbdata)),
    '<cacheSource type="worksheet"><worksheetSource ref="', ref, '" sheet="', sheet, '"/></cacheSource>',
    '<cacheFields count="', count, '">',
    paste0(cacheFields(wbdata, filter, rows, cols, data), collapse = ""),
    '</cacheFields>',
    '</pivotCacheDefinition>'
  )
}


pivot_rec_xml <- function(data) {
  paste0(
    sprintf('<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" count="%s">', nrow(data)),
    '</pivotCacheRecords>'
  )
}

pivot_def_rel <- function(n) sprintf("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" Target=\"pivotCacheRecords%s.xml\"/></Relationships>", n)

pivot_xml_rels <- function(n) sprintf("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition%s.xml\"/></Relationships>", n)

get_items <- function(data, x) {
  x <- abs(x)
  item <- sapply(
    c(order(unique(data[[x]])) - 1L, "default"),
    # # TODO this sets the order of the pivot elements
    # c(seq_along(unique(data[[x]])) - 1L, "default"),
    function(val) {
      if (val == "default")
        xml_node_create("item", xml_attributes = c(t = val))
      else
        xml_node_create("item", xml_attributes = c(x = val))
    },
    USE.NAMES = FALSE
  )
  items <- xml_node_create(
    "items",
    xml_attributes = c(count = as.character(length(item))), xml_children = item
  )
  items
}

row_col_items <- function(data, z, var) {
  var <- abs(var)
  item <- sapply(
    c(seq_along(unique(data[, var])) - 1L, "grand"),
    function(val) {
      if (val == "0") {
        xml_node_create("i", xml_children = xml_node_create("x"))
      } else if (val == "grand") {
        xml_node_create("i", xml_attributes = c(t = val), xml_children = xml_node_create("x"))
      } else {
        xml_node_create("i", xml_children = xml_node_create("x", xml_attributes = c(v = val)))
      }
    },
    USE.NAMES = FALSE
  )

  xml_node_create(z, xml_attributes = c(count = as.character(length(item))), xml_children = item)
}

create_pivot_table <- function(
    x,
    dims,
    filter,
    rows,
    cols,
    data,
    n,
    fun,
    params,
    numfmts
  ) {

  if (missing(filter)) {
    filter_pos <- 0
    use_filter <- FALSE
    # message("no filter")
  } else {
    filter_pos <- match(filter, names(x))
    use_filter <- TRUE
  }

  if (missing(rows)) {
    rows_pos <- 0
    use_rows <- FALSE
    # message("no rows")
  } else {
    rows_pos <- match(rows, names(x))
    use_rows <- TRUE
  }

  if (missing(cols)) {
    cols_pos <- 0
    use_cols <- FALSE
    # message("no cols")
  } else {
    cols_pos <- match(cols, names(x))
    use_cols <- TRUE
  }

  if (missing(data)) {
    data_pos <- 0
    use_data <- FALSE
    # message("no data")
  } else {
    data_pos <- match(data, names(x))
    use_data <- TRUE
    if (length(data_pos) > 1) {
      cols_pos <- c(-1L, cols_pos[cols_pos > 0])
      use_cols <- TRUE
    }
  }

  pivotField <- NULL
  for (i in seq_along(x)) {

    dataField <- NULL
    axis <- NULL
    sort <- NULL
    autoSortScope <- NULL

    if (i %in% data_pos)    dataField <- c(dataField = "1")

    if (i %in% filter_pos)  axis <- c(axis = "axisPage")

    if (i %in% rows_pos) {
      axis <- c(axis = "axisRow")
      sort <- params$sort_row

      if (!is.null(sort) && !is.character(sort)) {
        if (!abs(sort) %in% seq_along(rows_pos))
          warning("invalid sort position found")

       if (!abs(sort) == match(i, rows_pos))
        sort <- NULL
      }
    }

    if (i %in% cols_pos) {
      axis <- c(axis = "axisCol")
      sort <- params$sort_col

      if (!is.null(sort) && !is.character(sort)) {
        if (!abs(sort) %in% seq_along(cols_pos))
          warning("invalid sort position found")

       if (!abs(sort) == match(i, cols_pos))
        sort <- NULL
      }
    }

    if (!is.null(sort) && !is.character(sort)) {

      autoSortScope <- read_xml(sprintf('
        <autoSortScope>
          <pivotArea dataOnly="0" outline="0" fieldPosition="0">
            <references count="1">
            <reference field="4294967294" count="1" selected="0">
              <x v="%s" />
            </reference>
            </references>
          </pivotArea>
        </autoSortScope>
        ',
        abs(sort) - 1L), pointer = FALSE)

      if (sign(sort) == -1) sort <- "descending"
      else                  sort <- "ascending"
    }

    attrs <- c(axis, dataField, showAll = "0", sortType = sort)

    tmp <- xml_node_create(
      "pivotField",
      xml_attributes = attrs)

    if (i %in% c(filter_pos, rows_pos, cols_pos)) {
      tmp <- xml_node_create(
        "pivotField",
        xml_attributes = attrs,
        xml_children = paste0(paste0(get_items(x, i), collapse = ""), autoSortScope))
    }

    pivotField <- c(pivotField, tmp)
  }

  pivotFields <- xml_node_create(
    "pivotFields",
    xml_children = pivotField,
    xml_attributes = c(count = as.character(length(pivotField))))

  cols_field <- c(cols_pos - 1L)
  # if (length(data_pos) > 1)
  #   cols_field <- c(cols_field * -1, rep(1, length(data_pos) -1L))

  dim <- dims

  if (use_filter) {
    pageFields <- paste0(
      sprintf('<pageFields count="%s">', length(data_pos)),
      paste0(sprintf('<pageField fld="%s" hier="-1" />', filter_pos - 1L), collapse = ""),
      '</pageFields>')
  } else {
    pageFields <- ""
  }

  if (use_cols) {
    colsFields <- paste0(
      sprintf('<colFields count="%s">', length(cols_pos)),
      paste0(sprintf('<field x="%s" />', cols_field), collapse = ""),
      '</colFields>',
      paste0(row_col_items(data = x, "colItems", cols_pos), collapse = "")
    )
  } else {
    colsFields <- ""
  }

  if (use_rows) {
    rowsFields <- paste0(
      sprintf('<rowFields count="%s">', length(rows_pos)),
      paste0(sprintf('<field x="%s" />', rows_pos - 1L), collapse = ""),
      '</rowFields>',
      paste0(row_col_items(data = x, "rowItems", rows_pos), collapse = "")
    )
  } else {
    rowsFields <- ""
  }

  if (use_data) {

    dataField <- NULL
    for (i in seq_along(data)) {

      if (missing(fun)) fun <- NULL

      dataField <- c(
        dataField,
        xml_node_create(
          "dataField",
          xml_attributes = c(
            name      = sprintf("%s of %s", ifelse(is.null(fun[i]), "Sum", fun[i]), data[i]),
            fld       = sprintf("%s", data_pos[i] - 1L),
            subtotal  = fun[i],
            baseField = "0",
            baseItem  = "0",
            numFmtId  = numfmts[i]
          )
        )
      )
    }

    dataFields <- xml_node_create(
      "dataFields",
      xml_attributes = c(count = as.character(length(data_pos))),
      xml_children = paste0(dataField, collapse = "")
    )
  } else {
    dataFields <- ""
  }

  location <- xml_node_create(
    "location",
    xml_attributes = c(
      ref            = dim,
      firstHeaderRow = "1",
      firstDataRow   = "2",
      firstDataCol   = "1",
      rowPageCount   = as_binary(use_filter),
      colPageCount   = as_binary(use_filter)
    )
  )

  name <- "PivotStyleLight16"
  if (!is.null(params$name))
    name <- params$name

  dataCaption <- "Values"
  if (!is.null(params$dataCaption))
    dataCaption <- params$dataCaption

  showRowHeaders <- "1"
  if (!is.null(params$showRowHeaders))
    showRowHeaders <- params$showRowHeaders

  showColHeaders <- "1"
  if (!is.null(params$showColHeaders))
    showColHeaders <- params$showColHeaders

  showRowStripes <- "0"
  if (!is.null(params$showRowStripes))
    showRowStripes <- params$showRowStripes

  showColStripes <- "0"
  if (!is.null(params$showColStripes))
    showColStripes <- params$showColStripes

  showLastColumn <- "1"
  if (!is.null(params$showLastColumn))
    showLastColumn <- params$showLastColumn

  pivotTableStyleInfo <- xml_node_create(
    "pivotTableStyleInfo",
    xml_attributes = c(
      name           = name,
      showRowHeaders = showRowHeaders,
      showColHeaders = showColHeaders,
      showRowStripes = showRowStripes,
      showColStripes = showColStripes,
      showLastColumn = showLastColumn
    )
  )

  # extLst <- paste0(
  #   '<extLst>',
  #   '<ext xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout" uri="{747A6164-185A-40DC-8AA5-F01512510D54}">',
  #   '<xpdl:pivotTableDefinition16/>',
  #   '</ext>',
  #   '<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}">',
  #   '<x14:pivotCacheDefinition/>',
  #   '</ext>',
  #   '</extLst>'
  # )

  if (isTRUE(params$no_style))
    pivotTableStyleInfo <- ""

  indent <- "0"
  if (!is.null(params$indent))
    indent <- params$indent

  itemPrintTitles <- "1"
  if (!is.null(params$itemPrintTitles))
    itemPrintTitles <- params$itemPrintTitles

  multipleFieldFilters <- "0"
  if (!is.null(params$multipleFieldFilters))
    multipleFieldFilters <- params$multipleFieldFilters

  outline <- "1"
  if (!is.null(params$outline))
    outline <- params$outline

  outlineData <- "1"
  if (!is.null(params$outlineData))
    outlineData <- params$outlineData

  useAutoFormatting <- "1"
  if (!is.null(params$useAutoFormatting))
    useAutoFormatting <- params$useAutoFormatting

  applyAlignmentFormats <- "0"
  if (!is.null(params$applyAlignmentFormats))
    applyAlignmentFormats <- params$applyAlignmentFormats

  applyNumberFormats <- "0"
  if (!is.null(params$applyNumberFormats))
    applyNumberFormats <- params$applyNumberFormats

  applyBorderFormats <- "0"
  if (!is.null(params$applyBorderFormats))
    applyBorderFormats <- params$applyBorderFormats

  applyFontFormats <- "0"
  if (!is.null(params$applyFontFormats))
    applyFontFormats <- params$applyFontFormats

  applyPatternFormats <- "0"
  if (!is.null(params$applyPatternFormats))
    applyPatternFormats <- params$applyPatternFormats

  applyWidthHeightFormats <- "1"
  if (!is.null(params$applyWidthHeightFormats))
    applyWidthHeightFormats <- params$applyWidthHeightFormats

  xml_node_create(
    "pivotTableDefinition",
    xml_attributes = c(
      xmlns                   = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      `xmlns:mc`              = "http://schemas.openxmlformats.org/markup-compatibility/2006",
      `mc:Ignorable`          = "xr",
      `xmlns:xr`              = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
      name                    = sprintf("PivotTable%s", n),
      cacheId                 = as.character(n),
      applyNumberFormats      = applyNumberFormats,
      applyBorderFormats      = applyBorderFormats,
      applyFontFormats        = applyFontFormats,
      applyPatternFormats     = applyPatternFormats,
      applyAlignmentFormats   = applyAlignmentFormats,
      applyWidthHeightFormats = applyWidthHeightFormats,
      asteriskTotals          = params$asteriskTotals,
      autoFormatId            = params$autoFormatId,
      chartFormat             = params$chartFormat,
      dataCaption             = dataCaption,
      updatedVersion          = "8",
      minRefreshableVersion   = "3",
      useAutoFormatting       = useAutoFormatting,
      itemPrintTitles         = itemPrintTitles,
      createdVersion          = "8",
      indent                  = indent,
      outline                 = outline,
      outlineData             = outlineData,
      multipleFieldFilters    = multipleFieldFilters,
      colGrandTotals          = params$colGrandTotals,
      colHeaderCaption        = params$colHeaderCaption,
      compact                 = params$compact,
      compactData             = params$compactData,
      customListSort          = params$customListSort,
      dataOnRows              = params$dataOnRows,
      dataPosition            = params$dataPosition,
      disableFieldList        = params$disableFieldList,
      editData                = params$editData,
      enableDrill             = params$enableDrill,
      enableFieldProperties   = params$enableFieldProperties,
      enableWizard            = params$enableWizard,
      errorCaption            = params$errorCaption,
      fieldListSortAscending  = params$fieldListSortAscending,
      fieldPrintTitles        = params$fieldPrintTitles,
      grandTotalCaption       = params$grandTotalCaption,
      gridDropZones           = params$gridDropZones,
      immersive               = params$immersive,
      mdxSubqueries           = params$mdxSubqueries,
      missingCaption          = params$missingCaption,
      mergeItem               = params$mergeItem,
      pageOverThenDown        = params$pageOverThenDown,
      pageStyle               = params$pageStyle,
      pageWrap                = params$pageWrap,
      # pivotTableStyle
      preserveFormatting      = params$preserveFormatting,
      printDrill              = params$printDrill,
      published               = params$published,
      rowGrandTotals          = params$rowGrandTotals,
      rowHeaderCaption        = params$rowHeaderCaption,
      showCalcMbrs            = params$showCalcMbrs,
      showDataDropDown        = params$showDataDropDown,
      showDataTips            = params$showDataTips,
      showDrill               = params$showDrill,
      showDropZones           = params$showDropZones,
      showEmptyCol            = params$showEmptyCol,
      showEmptyRow            = params$showEmptyRow,
      showError               = params$showError,
      showHeaders             = params$showHeaders,
      showItems               = params$showItems,
      showMemberPropertyTips  = params$showMemberPropertyTips,
      showMissing             = params$showMissing,
      showMultipleLabel       = params$showMultipleLabel,
      subtotalHiddenItems     = params$subtotalHiddenItems,
      tag                     = params$tag,
      vacatedStyle            = params$vacatedStyle,
      visualTotals            = params$visualTotals
    ),
    # xr:uid="{375073AB-E7CA-C149-922E-A999C47476C1}"
    xml_children =  paste0(
      location,
      paste0(pivotFields, collapse = ""),
      rowsFields,
      colsFields,
      pageFields,
      dataFields,
      pivotTableStyleInfo
      # extLst
    )
  )

}

### end pivot table helpers

### modify xml file names

read_Content_Types <- function(x) {

  df <- rbindlist(xml_attr(x, "Default"))
  or <- rbindlist(xml_attr(x, "Override"))

  sel <- grepl("/sheet[0-9]+.xml$", or$PartName)
  or$sheet_id <- NA
  or$sheet_id[sel] <- as.integer(gsub("\\D+", "", basename(or$PartName)[sel]))


  list(df, or)
}

write_Content_Types <- function(x, rm_sheet = NULL) {
  df_df <- x[[1]]
  or_df <- x[[2]]

  if (!is.null(rm_sheet)) {
    # remove a sheet and rename the remaining
    or_df <- or_df[order(or_df$sheet_id), ]
    sel <- which(or_df$sheet_id == rm_sheet)
    or_df <- or_df[-sel, ]
    or_df$PartName[!is.na(or_df$sheet_id)] <- sprintf(
      "%s/sheet%s.xml",
      dirname(or_df$PartName[!is.na(or_df$sheet_id)]),
      seq_along(or_df$sheet_id[!is.na(or_df$sheet_id)])
    )

    or_df <- or_df[order(as.integer(row.names(or_df))), ]
  }

  df_xml <- df_to_xml("Default", df_df[c("Extension", "ContentType")])
  or_xml <- df_to_xml("Override", or_df[c("PartName", "ContentType")])
  c(df_xml, or_xml)
}

read_workbook.xml.rels <- function(x) {

  wxr <- rbindlist(xml_attr(x, "Relationship"))

  sel <- grepl("/sheet[0-9]+.xml$", wxr$Target)
  wxr$sheet_id <- NA
  wxr$sheet_id[sel] <- as.integer(gsub("\\D+", "", basename(wxr$Target)[sel]))
  wxr
}

write_workbook.xml.rels <- function(x, rm_sheet = NULL) {
  wxr <- x

  if (!is.null(rm_sheet)) {
    # remove a sheet and rename the remaining
    wxr <- wxr[order(wxr$sheet_id), ]
    sel <- which(wxr$sheet_id == rm_sheet)
    wxr <- wxr[-sel, ]
    wxr$Target[!is.na(wxr$sheet_id)] <- sprintf(
      "%s/sheet%s.xml",
      dirname(wxr$Target[!is.na(wxr$sheet_id)]),
      seq_along(wxr$sheet_id[!is.na(wxr$sheet_id)])
    )
  }

  if (is.null(wxr[["TargetMode"]])) wxr$TargetMode <- ""
  df_to_xml("Relationship", df_col = wxr[c("Id", "Type", "Target", "TargetMode")])
}

#' convert objects with attribute labels into strings
#' @param x an object to convert
#' @keywords internal
#' @noRd
to_string <- function(x) {
  lbls <- attr(x, "labels")
  chr <- as.character(x)
  if (!is.null(lbls)) {
    lbls <- lbls[lbls %in% x]
    sel_l <- match(lbls, x)
    if (length(sel_l)) chr[sel_l] <- names(lbls)
  }
  chr
}
