# `create_hyperlink()` ---------------------------------------------------------
#' Create Excel hyperlink string
#'
#' Wrapper to create internal hyperlink string to pass to [write_formula()].
#' Either link to external URLs or local files or straight to cells of local Excel sheets.
#'
#' @param sheet Name of a worksheet
#' @param row integer row number for hyperlink to link to
#' @param col column number of letter for hyperlink to link to
#' @param text display text
#' @param file Excel file name to point to. If NULL hyperlink is internal.
#' @examples
#' wb <- wb_workbook()$
#'   add_worksheet("Sheet1")$add_worksheet("Sheet2")$add_worksheet("Sheet3")
#'
#' ## Internal Hyperlink - create hyperlink formula manually
#' x <- '=HYPERLINK(\"#Sheet2!B3\", "Text to Display - Link to Sheet2")'
#' wb$add_formula(sheet = "Sheet1", x = x, dims = "A1")
#'
#' ## Internal - No text to display using create_hyperlink() function
#' x <- create_hyperlink(sheet = "Sheet3", row = 1, col = 2)
#' wb$add_formula(sheet = "Sheet1", x = x, dims = "A2")
#'
#' ## Internal - Text to display
#' x <- create_hyperlink(sheet = "Sheet3", row = 1, col = 2,text = "Link to Sheet 3")
#' wb$add_formula(sheet = "Sheet1", x = x, dims = "A3")
#'
#' ## Link to file - No text to display
#' fl <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' x <- create_hyperlink(sheet = "Sheet1", row = 3, col = 10, file = fl)
#' wb$add_formula(sheet = "Sheet1", x = x, dims = "A4")
#'
#' ## Link to file - Text to display
#' fl <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' x <- create_hyperlink(sheet = "Sheet2", row = 3, col = 10, file = fl, text = "Link to File.")
#' wb$add_formula(sheet = "Sheet1", x = x, dims = "A5")
#'
#' ## Link to external file - Text to display
#' x <- '=HYPERLINK("[C:/Users]", "Link to an external file")'
#' wb$add_formula(sheet = "Sheet1", x = x, dims = "A6")
#'
#' x <- create_hyperlink(text = "test.png", file = "D:/somepath/somepicture.png")
#' wb$add_formula(x = x, dims = "A7")
#'
#' # if (interactive()) wb$open()
#'
#' @export
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

getId <- function(x) reg_match0(x, '(?<= Id=")[0-9A-Za-z]+')

# `col2hex()` -----------------------------------------------------------------
#' Convert rgb to hex
#'
#' @param my.col my.col
#' @noRd
col2hex <- function(my.col) {
  rgb(t(col2rgb(my.col)), maxColorValue = 255)
}

# validate color --------------------------------------------------------------
#' Validate and convert color. Returns ARGB string
#' @param color input string (something like color(), "00000", "#000000", "00000000" or "#00000000")
#' @param or_null logical for use in assert
#' @param envir parent frame for use in assert
#' @param msg return message
#' @noRd
validate_color <- function(color = NULL, or_null = FALSE, envir = parent.frame(), msg = NULL) {
  sx <- as.character(substitute(color, envir))

  if (identical(color, "none") && or_null) {
    return(NULL)
  }

  # returns black
  if (is.null(color)) {
    if (or_null) return(NULL)
    return("FF000000")
  }

  if (any(ind <- color %in% grDevices::colors())) {
    color[ind] <- col2hex(color[ind])
  }

  # remove any # from color strings
  color <- gsub("^#", "", toupper(color))

  ## create a total size of 8 in ARGB format
  color <- stringi::stri_pad_left(str = color, width = 8, pad = "F")

  if (any(!grepl("[A-F0-9]{8}$", color))) {
    if (is.null(msg)) msg <- sprintf("`%s` ['%s'] is not a valid color", sx, color)
    stop(simpleError(msg))
  }

  color
}

## header and footer replacements ----------------------------------------------
headerFooterSub <- function(x) {
  if (!is.null(x)) {
    x <- replace_legal_chars(x)
    x <- gsub("\\[Page\\]", "P", x)
    x <- gsub("\\[Pages\\]", "N", x)
    x <- gsub("\\[Date\\]", "D", x)
    x <- gsub("\\[Time\\]", "T", x)
    x <- gsub("\\[Path\\]", "Z", x)
    x <- gsub("\\[File\\]", "F", x)
    x <- gsub("\\[Tab\\]", "A", x)
  }

  x
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

      # either a <r> node or <t> from unstyled comment
      is_fmt_txt <- FALSE
      if (is_xml(comment))
        is_fmt_txt <- all(xml_node_name(comment) == "r") || isFALSE(comment_list[[i]]$style[[j]])

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

# `amp_split()` ----------------------------------------------------------------
#' split headerFooter xml into left, center, and right.
#' @param x xml string
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

# Header footer ---------------------------------------------------------------
#' get headerFooter from xml into list with left, center, and right.
#' @param x xml string
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

#' Create sparklines object
#'
#' Create a sparkline to be added a workbook with [wb_add_sparklines()]
#'
#' Colors are all predefined to be rgb. Maybe theme colors can be
#' used too.
#' @param sheet sheet
#' @param dims Cell range of cells used to create the sparklines
#' @param sqref Cell range of the destination of the sparklines.
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
#' @return A string containing XML code
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
  # TODO change arguments to snake case?

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
      "xr2:uid" = sprintf("{6F57B887-24F1-C14A-942C-%s}", random_string(length = 12, pattern = "[A-F0-9]"))
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


# unique is not enough! gh#713
distinct <- function(x) {
  unis <- stringi::stri_unique(x)
  lwrs <- tolower(unis)
  dups <- duplicated(lwrs)
  unis[dups == FALSE]
}

cacheFields <- function(wbdata, filter, rows, cols, data, slicer) {
  sapply(
    names(wbdata),
    function(x) {

      dat <- wbdata[[x]]

      vars <- c(filter, rows, cols, data, slicer)

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

      if (is_vars && !is_data) {
        sharedItem <- vapply(
          distinct(dat),
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

pivot_def_xml <- function(wbdata, filter, rows, cols, data, slicer, pcid) {

  ref   <- dataframe_to_dims(attr(wbdata, "dims"))
  sheet <- attr(wbdata, "sheet")
  count <- ncol(wbdata)

  paste0(
    sprintf('<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="xr" r:id="rId1" refreshedBy="openxlsx2" invalid="1" refreshOnLoad="1" refreshedDate="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="%s">', nrow(wbdata)),
    '<cacheSource type="worksheet"><worksheetSource ref="', ref, '" sheet="', sheet, '"/></cacheSource>',
    '<cacheFields count="', count, '">',
    paste0(cacheFields(wbdata, filter, rows, cols, data, slicer), collapse = ""),
    '</cacheFields>',
    '<extLst>',
    '<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}">',
    sprintf('<x14:pivotCacheDefinition pivotCacheId="%s"/>', pcid),
    '</ext>',
    '</extLst>',
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

get_items <- function(data, x, item_order, slicer = FALSE, choose = NULL) {
  x <- abs(x)

  dat <- distinct(data[[x]])

  # check length, otherwise a certain spreadsheet software simply dies
  if (!is.null(item_order) && (length(item_order) != length(dat))) {
    msg <- sprintf(
      "Length of sort order for '%s' does not match required length. Is %s, needs %s.\nCheck `openxlsx2:::distinct()` for the correct length. Resetting.",
      names(data[x]), length(item_order), length(dat)
    )
    warning(msg)
    item_order <- NULL
  }

  if (is.null(item_order)) {
    item_order <- order(dat)
  } else if (is.character(item_order)) {
    item_order <- match(dat, item_order)
  }

  if (!is.null(choose)) {
    # change order
    choose <- eval(parse(text = choose), data.frame(x = dat))[item_order]
    hide <- as_xml_attr(!choose)
    sele <- as_xml_attr(choose)
  } else {
    hide <- NULL
    sele <- rep("1", length(dat))
  }

  if (slicer) {
    vals <- as.character(item_order - 1L)
    item <- sapply(
      seq_along(vals),
      function(val) {
          xml_node_create("i", xml_attributes = c(x = vals[val], s = sele[val]))
      },
      USE.NAMES = FALSE
    )
  } else {
    vals <- c(item_order - 1L, "default")
    item <- sapply(
      seq_along(vals),
      # # TODO this sets the order of the pivot elements
      # c(seq_along(unique(data[[x]])) - 1L, "default"),
      function(val) {
        if (vals[val] == "default")
          xml_node_create("item", xml_attributes = c(t = vals[val]))
        else
          xml_node_create("item", xml_attributes = c(x = vals[val], h = hide[val]))
      },
      USE.NAMES = FALSE
    )
  }

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
  arguments <- c(
    "apply_alignment_formats", "apply_border_formats", "apply_font_formats",
    "apply_number_formats", "apply_pattern_formats", "apply_width_height_formats",
    "apply_width_height_formats", "asterisk_totals", "auto_format_id",
    "chart_format", "col_grand_totals", "col_header_caption", "compact",
    "compact", "choose", "compact_data", "custom_list_sort", "data_caption",
    "data_on_rows", "data_position", "disable_field_list", "edit_data",
    "enable_drill", "enable_field_properties", "enable_wizard", "error_caption",
    "field_list_sort_ascending", "field_print_titles", "grand_total_caption",
    "grid_drop_zones", "immersive", "indent", "item_print_titles",
    "mdx_subqueries", "merge_item", "missing_caption", "multiple_field_filters",
    "name", "no_style", "numfmt", "outline", "outline_data", "page_over_then_down",
    "page_style", "page_wrap", "preserve_formatting", "print_drill",
    "published", "row_grand_totals", "row_header_caption", "show_calc_mbrs",
    "show_col_headers", "show_col_stripes", "show_data_as", "show_data_drop_down",
    "show_data_tips", "show_drill", "show_drop_zones", "show_empty_col",
    "show_empty_row", "show_error", "show_headers", "show_items",
    "show_last_column", "show_member_property_tips", "show_missing",
    "show_multiple_label", "show_row_headers", "show_row_stripes",
    "sort_col", "sort_item", "sort_row", "subtotal_hidden_items",
    "table_style", "tag", "use_auto_formatting", "vacated_style",
    "visual_totals"
  )
  params <- standardize_case_names(params, arguments = arguments, return = TRUE)


  compact <- ""
  if (!is.null(params$compact))
    compact <- params$compact

  outline <- ""
  if (!is.null(params$outline))
    outline <- params$outline

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

    sort_item <- params$sort_item
    choose    <- params$choose
    multi     <- if (is.null(choose)) NULL else as_xml_attr(TRUE)

    attrs <- c(
      axis, dataField, showAll = "0", multipleItemSelectionAllowed = multi, sortType = sort,
      compact = as_xml_attr(compact), outline = as_xml_attr(outline)
    )

    tmp <- xml_node_create(
      "pivotField",
      xml_attributes = attrs)

    if (i %in% c(filter_pos, rows_pos, cols_pos)) {
      nms <- names(x[i])
      sort_itm <- sort_item[[nms]]
      if (!is.null(choose) && !is.na(choose[nms])) {
        choo <- choose[nms]
      } else {
        choo <- NULL
      }
      tmp <- xml_node_create(
        "pivotField",
        xml_attributes = attrs,
        xml_children = paste0(paste0(get_items(x, i, sort_itm, FALSE, choo), collapse = ""), autoSortScope))
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

    show_data_as <- rep("", length(data))
    if (!is.null(params$show_data_as))
      show_data_as <- params$show_data_as

    for (i in seq_along(data)) {

      if (missing(fun)) fun <- NULL

      dataField <- c(
        dataField,
        xml_node_create(
          "dataField",
          xml_attributes = c(
            name       = sprintf("%s of %s", ifelse(is.null(fun[i]), "Sum", fun[i]), data[i]),
            fld        = sprintf("%s", data_pos[i] - 1L),
            showDataAs = as_xml_attr(ifelse(is.null(show_data_as[i]), "", show_data_as[i])),
            subtotal   = fun[i],
            baseField  = "0",
            baseItem   = "0",
            numFmtId   = numfmts[i]
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

  table_style <- "PivotStyleLight16"
  if (!is.null(params$table_style))
    table_style <- params$table_style

  dataCaption <- "Values"
  if (!is.null(params$data_caption))
    dataCaption <- params$data_caption

  showRowHeaders <- "1"
  if (!is.null(params$show_row_headers))
    showRowHeaders <- params$show_row_headers

  showColHeaders <- "1"
  if (!is.null(params$show_col_headers))
    showColHeaders <- params$show_col_headers

  showRowStripes <- "0"
  if (!is.null(params$show_row_stripes))
    showRowStripes <- params$show_row_stripes

  showColStripes <- "0"
  if (!is.null(params$show_col_stripes))
    showColStripes <- params$show_col_stripes

  showLastColumn <- "1"
  if (!is.null(params$show_last_column))
    showLastColumn <- params$show_last_column

  pivotTableStyleInfo <- xml_node_create(
    "pivotTableStyleInfo",
    xml_attributes = c(
      name           = table_style,
      showRowHeaders = showRowHeaders,
      showColHeaders = showColHeaders,
      showRowStripes = showRowStripes,
      showColStripes = showColStripes,
      showLastColumn = showLastColumn
    )
  )

  if (isTRUE(params$no_style))
    pivotTableStyleInfo <- ""

  indent <- "0"
  if (!is.null(params$indent))
    indent <- params$indent

  itemPrintTitles <- "1"
  if (!is.null(params$item_print_titles))
    itemPrintTitles <- params$item_print_titles

  multipleFieldFilters <- "0"
  if (!is.null(params$multiple_field_filters))
    multipleFieldFilters <- params$multiple_field_filters

  outlineData <- "1"
  if (!is.null(params$outline_data))
    outlineData <- params$outline_data

  useAutoFormatting <- "1"
  if (!is.null(params$use_auto_formatting))
    useAutoFormatting <- params$use_auto_formatting

  applyAlignmentFormats <- "0"
  if (!is.null(params$apply_alignment_formats))
    applyAlignmentFormats <- params$apply_alignment_formats

  applyNumberFormats <- "0"
  if (!is.null(params$apply_number_formats))
    applyNumberFormats <- params$apply_number_formats

  applyBorderFormats <- "0"
  if (!is.null(params$apply_border_formats))
    applyBorderFormats <- params$apply_border_formats

  applyFontFormats <- "0"
  if (!is.null(params$apply_font_formats))
    applyFontFormats <- params$apply_font_formats

  applyPatternFormats <- "0"
  if (!is.null(params$apply_pattern_formats))
    applyPatternFormats <- params$apply_pattern_formats

  applyWidthHeightFormats <- "1"
  if (!is.null(params$apply_width_height_formats))
    applyWidthHeightFormats <- params$apply_width_height_formats

  pivot_table_name <- sprintf("PivotTable%s", n)
  if (!is.null(params$name))
    pivot_table_name <- params$name

  xml_node_create(
    "pivotTableDefinition",
    xml_attributes = c(
      xmlns                   = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      `xmlns:mc`              = "http://schemas.openxmlformats.org/markup-compatibility/2006",
      `mc:Ignorable`          = "xr",
      `xmlns:xr`              = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
      name                    = as_xml_attr(pivot_table_name),
      cacheId                 = as_xml_attr(n),
      applyNumberFormats      = as_xml_attr(applyNumberFormats),
      applyBorderFormats      = as_xml_attr(applyBorderFormats),
      applyFontFormats        = as_xml_attr(applyFontFormats),
      applyPatternFormats     = as_xml_attr(applyPatternFormats),
      applyAlignmentFormats   = as_xml_attr(applyAlignmentFormats),
      applyWidthHeightFormats = as_xml_attr(applyWidthHeightFormats),
      asteriskTotals          = as_xml_attr(params$asterisk_totals),
      autoFormatId            = as_xml_attr(params$auto_format_id),
      chartFormat             = as_xml_attr(params$chart_format),
      dataCaption             = as_xml_attr(dataCaption),
      updatedVersion          = "8",
      minRefreshableVersion   = "3",
      useAutoFormatting       = as_xml_attr(useAutoFormatting),
      itemPrintTitles         = as_xml_attr(itemPrintTitles),
      createdVersion          = "8",
      indent                  = as_xml_attr(indent),
      outline                 = as_xml_attr(outline),
      outlineData             = as_xml_attr(outlineData),
      multipleFieldFilters    = as_xml_attr(multipleFieldFilters),
      colGrandTotals          = as_xml_attr(params$col_grand_totals),
      colHeaderCaption        = as_xml_attr(params$col_header_caption),
      compact                 = as_xml_attr(params$compact),
      compactData             = as_xml_attr(params$compact_data),
      customListSort          = as_xml_attr(params$custom_list_sort),
      dataOnRows              = as_xml_attr(params$data_on_rows),
      dataPosition            = as_xml_attr(params$data_position),
      disableFieldList        = as_xml_attr(params$disable_field_list),
      editData                = as_xml_attr(params$edit_data),
      enableDrill             = as_xml_attr(params$enable_drill),
      enableFieldProperties   = as_xml_attr(params$enable_field_properties),
      enableWizard            = as_xml_attr(params$enable_wizard),
      errorCaption            = as_xml_attr(params$error_caption),
      fieldListSortAscending  = as_xml_attr(params$field_list_sort_ascending),
      fieldPrintTitles        = as_xml_attr(params$field_print_titles),
      grandTotalCaption       = as_xml_attr(params$grand_total_caption),
      gridDropZones           = as_xml_attr(params$grid_drop_zones),
      immersive               = as_xml_attr(params$immersive),
      mdxSubqueries           = as_xml_attr(params$mdx_subqueries),
      missingCaption          = as_xml_attr(params$missing_caption),
      mergeItem               = as_xml_attr(params$merge_item),
      pageOverThenDown        = as_xml_attr(params$page_over_then_down),
      pageStyle               = as_xml_attr(params$page_style),
      pageWrap                = as_xml_attr(params$page_wrap),
      # pivotTableStyle
      preserveFormatting      = as_xml_attr(params$preserve_formatting),
      printDrill              = as_xml_attr(params$print_drill),
      published               = as_xml_attr(params$published),
      rowGrandTotals          = as_xml_attr(params$row_grand_totals),
      rowHeaderCaption        = as_xml_attr(params$row_header_caption),
      showCalcMbrs            = as_xml_attr(params$show_calc_mbrs),
      showDataDropDown        = as_xml_attr(params$show_data_drop_down),
      showDataTips            = as_xml_attr(params$show_data_tips),
      showDrill               = as_xml_attr(params$show_drill),
      showDropZones           = as_xml_attr(params$show_drop_zones),
      showEmptyCol            = as_xml_attr(params$show_empty_col),
      showEmptyRow            = as_xml_attr(params$show_empty_row),
      showError               = as_xml_attr(params$show_error),
      showHeaders             = as_xml_attr(params$show_headers),
      showItems               = as_xml_attr(params$show_items),
      showMemberPropertyTips  = as_xml_attr(params$show_member_property_tips),
      showMissing             = as_xml_attr(params$show_missing),
      showMultipleLabel       = as_xml_attr(params$show_multiple_label),
      subtotalHiddenItems     = as_xml_attr(params$subtotal_hidden_items),
      tag                     = as_xml_attr(params$tag),
      vacatedStyle            = as_xml_attr(params$vacated_style),
      visualTotals            = as_xml_attr(params$visual_totals)
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

# get the next free relationship id
get_next_id <- function(x) {
  if (length(x)) {
    rlshp <- rbindlist(xml_attr(x, "Relationship"))
    rlshp$id <- as.integer(gsub("\\D+", "", rlshp$Id))
    next_id <- paste0("rId", max(rlshp$id) + 1L)
  } else {
    next_id <- "rId1"
  }
  next_id
}

#' create a guid string
#' @keywords internal
#' @noRd
st_guid <- function() {
  paste0(
    "{",
    random_string(length = 8, pattern = "[A-F0-9]"), "-",
    random_string(length = 4, pattern = "[A-F0-9]"), "-",
    random_string(length = 4, pattern = "[A-F0-9]"), "-",
    random_string(length = 4, pattern = "[A-F0-9]"), "-",
    random_string(length = 12, pattern = "[A-F0-9]"),
    "}"
  )
}

#' create a userid
#' @keywords internal
#' @noRd
st_userid <- function() {
  random_string(length = 16, pattern = "[a-z0-9]")
}

# solve merge helpers -----------------------------------------------------

#' check side
#' @param x a logical string
#' @name sidehelper
#' @noRd
fullsided <- function(x) {
  x[1] && x[length(x)]
}

#' @rdname sidehelper
#' @noRd
onesided <- function(x) {
  (x[1] && !x[length(x)]) || (!x[1] && x[length(x)])
}


#' @rdname sidehelper
#' @noRd
twosided <- function(x) {
  if (any(x)) (!x[1] && !x[length(x)])
  else FALSE
}

#' @rdname sidehelper
#' @noRd
top_half <- function(x) {
  if (twosided(x)) {
    out <- rep(FALSE, length(x))
    out[seq_len(which(x == TRUE)[1] - 1L)] <- TRUE
    return(out)
  } else {
    stop("not twosided")
  }
}

#' @rdname sidehelper
#' @noRd
bottom_half <- function(x) {
  if (twosided(x)) {
    out <- rep(TRUE, length(x))
    out[seq_len(rev(which(x == TRUE))[1])] <- FALSE
    return(out)
  } else {
    stop("not twosided")
  }
}

#' merge solver. split exisisting merge into pieces
#' @param have current merged cells
#' @param want newly merged cells
#' @noRd
solve_merge <- function(have, want) {

  got <- dims_to_dataframe(have, fill = TRUE)
  new <- dims_to_dataframe(want, fill = TRUE)

  cols_overlap <- colnames(got) %in% colnames(new)
  rows_overlap <- rownames(got) %in% rownames(new)

  # no overlap at all
  if (!any(cols_overlap) || !any(rows_overlap)) {
    return(have)
  }

  # return pieces of the old
  pieces <- list()

  # new overlaps old completely
  if (fullsided(cols_overlap) && fullsided(rows_overlap)) {
    return(NA_character_)
  }

  # all columns are overlapped onesided
  if (fullsided(cols_overlap) && onesided(rows_overlap)) {
    pieces[[1]] <- got[!rows_overlap, drop = FALSE]
  }

  # all columns are overlapped twosided
  if (fullsided(cols_overlap) && twosided(rows_overlap)) {
    pieces[[1]] <- got[top_half(rows_overlap), drop = FALSE]
    pieces[[2]] <- got[bottom_half(rows_overlap), drop = FALSE]
  }

  # all rows are overlapped onesided
  if (onesided(cols_overlap) && fullsided(rows_overlap)) {
    pieces[[1]] <- got[, !cols_overlap, drop = FALSE]
  }

  # all rows are overlapped twosided
  if (twosided(cols_overlap) && fullsided(rows_overlap)) {
    pieces[[1]] <- got[, top_half(cols_overlap), drop = FALSE]
    pieces[[2]] <- got[, bottom_half(cols_overlap), drop = FALSE]
  }

  # new is part of old
  if (onesided(cols_overlap) && onesided(rows_overlap)) {
    pieces[[1]] <- got[!rows_overlap, cols_overlap, drop = FALSE]
    pieces[[2]] <- got[, !cols_overlap, drop = FALSE]
  }

  if (onesided(cols_overlap) && twosided(rows_overlap)) {
    pieces[[1]] <- got[top_half(rows_overlap), cols_overlap, drop = FALSE]
    pieces[[2]] <- got[bottom_half(rows_overlap), cols_overlap, drop = FALSE]
    pieces[[3]] <- got[, !cols_overlap, drop = FALSE]
  }

  if (twosided(cols_overlap) && onesided(rows_overlap)) {
    pieces[[1]] <- got[rows_overlap, top_half(cols_overlap), drop = FALSE]
    pieces[[2]] <- got[rows_overlap, bottom_half(cols_overlap), drop = FALSE]
    pieces[[3]] <- got[!rows_overlap, , drop = FALSE]
  }

  if (twosided(cols_overlap) && twosided(rows_overlap)) {
    pieces[[1]] <- got[, top_half(cols_overlap), drop = FALSE]
    pieces[[2]] <- got[, bottom_half(cols_overlap), drop = FALSE]
    pieces[[3]] <- got[top_half(rows_overlap), cols_overlap, drop = FALSE]
    pieces[[4]] <- got[bottom_half(rows_overlap), cols_overlap, drop = FALSE]
  }

  vapply(pieces, dataframe_to_dims, NA_character_)
}

#' get the basename
#' on windows [basename()] only handles strings up to 255 characters, but we
#' can have longer strings when loading file systems
#' @param path a character string
#' @keywords internal
#' @noRd
basename2 <- function(path) {
  is_to_long <- vapply(path, to_long, NA)
  if (any(is_to_long)) {
    return(gsub(".*[\\/]", "", path))
  } else {
    return(basename(path))
  }
}

fetch_styles <- function(wb, xf_xml, st_ids) {
  # returns NA if no style found
  if (all(is.na(xf_xml))) return(NULL)

  lst_out <- vector("list", length = length(xf_xml))
  names(lst_out) <- st_ids

  for (i in seq_along(xf_xml)) {

    if (is.na(xf_xml[[i]])) next
    xf_df <- read_xf(read_xml(xf_xml[[i]]))

    border_id <- which(wb$styles_mgr$border$id == xf_df$borderId)
    fill_id   <- which(wb$styles_mgr$fill$id == xf_df$fillId)
    font_id   <- which(wb$styles_mgr$font$id == xf_df$fontId)
    numFmt_id <- which(wb$styles_mgr$numfmt$id == xf_df$numFmtId)

    border_xml <- wb$styles_mgr$styles$borders[border_id]
    fill_xml   <- wb$styles_mgr$styles$fills[fill_id]
    font_xml   <- wb$styles_mgr$styles$fonts[font_id]
    numfmt_xml <- wb$styles_mgr$styles$numFmts[numFmt_id]

    out <- list(
      xf_df,
      border_xml,
      fill_xml,
      font_xml,
      numfmt_xml
    )
    names(out) <- c("xf_df", "border_xml", "fill_xml", "font_xml", "numfmt_xml")
    lst_out[[i]] <- out

  }

  # unique drops names
  lst_out <- lst_out[!duplicated(lst_out)]

  attr(lst_out, "st_ids") <- st_ids

  lst_out
}

## get cell styles for a worksheet
get_cellstyle <- function(wb, sheet = current_sheet(), dims) {

  st_ids <- NULL
  if (missing(dims)) {
    st_ids <- styles_on_sheet(wb = wb, sheet = sheet) %>% as.character()
    xf_ids <- match(st_ids, wb$styles_mgr$xf$id)
    xf_xml <- wb$styles_mgr$styles$cellXfs[xf_ids]
  } else {
    xf_xml <- get_cell_styles(wb = wb, sheet = sheet, cell = dims)
  }

  fetch_styles(wb, xf_xml, st_ids)
}

get_colstyle <- function(wb, sheet = current_sheet()) {

  st_ids <- NULL
  if (length(wb$worksheets[[sheet]]$cols_attr)) {
    cols <- wb$worksheets[[sheet]]$unfold_cols()
    st_ids <- cols$s[cols$s != ""]
    xf_ids <- match(st_ids, wb$styles_mgr$xf$id)
    xf_xml <- wb$styles_mgr$styles$cellXfs[xf_ids]
  } else {
    xf_xml <- NA_character_
  }

  fetch_styles(wb, xf_xml, st_ids)
}

get_rowstyle <- function(wb, sheet = current_sheet()) {

  st_ids <- NULL
  if (!is.null(wb$worksheets[[sheet]]$sheet_data$row_attr)) {
    rows <- wb$worksheets[[sheet]]$sheet_data$row_attr
    st_ids <- rows$s[rows$s != ""]
    xf_ids <- match(st_ids, wb$styles_mgr$xf$id)
    xf_xml <- wb$styles_mgr$styles$cellXfs[xf_ids]
  } else {
    xf_xml <- NA_character_
  }

  fetch_styles(wb, xf_xml, st_ids)
}

## apply cell styles to a worksheet and return reference ids
set_cellstyles <- function(wb, style) {

  session_ids <- random_string(n = length(style))

  for (i in seq_along(style)) {
    session_id <- session_ids[i]

    has_border <- FALSE
    if (length(style[[i]]$border_xml)) {
      has_border <- TRUE
      wb$styles_mgr$add(style[[i]]$border_xml, session_id)
    }

    has_fill <- FALSE
    if (length(style[[i]]$fill_xml) && style[[i]]$fill_xml != wb$styles_mgr$styles$fills[1]) {
      has_fill <- TRUE
      wb$styles_mgr$add(style[[i]]$fill_xml, session_id)
    }

    has_font <- FALSE
    if (length(style[[i]]$font_xml) && style[[i]]$font_xml != wb$styles_mgr$styles$fonts[1]) {
      has_font <- TRUE
      wb$styles_mgr$add(style[[i]]$font_xml, session_id)
    }

    has_numfmt <- FALSE
    if (length(style[[i]]$numfmt_xml)) {
      has_numfmt <- TRUE
      numfmt_xml <- style[[i]]$numfmt_xml
      # assuming all numfmts with ids >= 164.
      # We have to create unique numfmt ids when cloning numfmts. Otherwise one
      # ids could point to more than one format code and the output would look
      # broken.
      fmtCode <- xml_attr(numfmt_xml, "numFmt")[[1]][["formatCode"]]
      next_id <- max(163L, as.integer(wb$styles_mgr$get_numfmt()$id)) + 1L
      numfmt_xml <- create_numfmt(numFmtId = next_id, formatCode = fmtCode)

      wb$styles_mgr$add(numfmt_xml, session_id)
    }

    ## create new xf_df. This has to reference updated style ids
    xf_df <- style[[i]]$xf_df

    if (has_border)
      xf_df$borderId <- wb$styles_mgr$get_border_id(session_id)

    if (has_fill)
      xf_df$fillId <- wb$styles_mgr$get_fill_id(session_id)

    if (has_font)
      xf_df$fontId <- wb$styles_mgr$get_font_id(session_id)

    if (has_numfmt)
      xf_df$numFmtId <- wb$styles_mgr$get_numfmt_id(session_id)

    xf_xml <- write_xf(xf_df) # can be NULL

    if (length(xf_xml))
      wb$styles_mgr$add(xf_xml, session_id)
  }

  # return updated style id
  st_ids <- wb$styles_mgr$get_xf_id(session_ids)

  if (!is.null(attr(style, "st_ids"))) {
    names(st_ids) <- names(style)
    out <- attr(style, "st_ids")

    want <- match(out, names(st_ids))
    st_ids <- st_ids[want]
  }

  st_ids
}

clone_shared_strings <- function(wb_old, old, wb_new, new) {

  empty <- structure(list(), uniqueCount = 0)

  # old has no shared strings
  if (identical(wb_old$sharedStrings, empty)) {
    return(NULL)
  }

  if (identical(wb_new$sharedStrings, empty)) {

    wb_new$append(
      "Content_Types",
      "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
    )

    wb_new$append(
      "workbook.xml.rels",
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
    )

  }

  sheet_id <- wb_old$validate_sheet(old)
  cc <- wb_old$worksheets[[sheet_id]]$sheet_data$cc
  sst_ids  <- as.integer(cc$v[cc$c_t == "s"]) + 1
  sst_uni  <- sort(unique(sst_ids))
  sst_old <- wb_old$sharedStrings[sst_uni]

  old_len <- length(as.character(wb_new$sharedStrings))

  wb_new$sharedStrings <- c(as.character(wb_new$sharedStrings), sst_old)
  sst  <- xml_node_create("sst", xml_children = wb_new$sharedStrings)
  text <- xml_si_to_txt(read_xml(sst))
  attr(wb_new$sharedStrings, "uniqueCount") <- as.character(length(text))
  attr(wb_new$sharedStrings, "text") <- text


  sheet_id <- wb_new$validate_sheet(new)
  cc <- wb_new$worksheets[[sheet_id]]$sheet_data$cc
  # order ids and add new offset
  ids <- as.integer(cc$v[cc$c_t == "s"]) + 1L
  new_ids <- match(ids, sst_uni) + old_len - 1L
  new_ids <- as.character(new_ids)
  new_ids[is.na(new_ids)] <- ""
  cc$v[cc$c_t == "s"] <- new_ids
  wb_new$worksheets[[sheet_id]]$sheet_data$cc <- cc

  # print(sprintf("cloned: %s", length(new_ids)))

}
