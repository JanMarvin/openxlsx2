# `create_hyperlink()` ---------------------------------------------------------
#' Create Excel hyperlink string
#'
#' @description
#' Wrapper to create internal hyperlink string to pass to [wb_add_formula()].
#' Either link to external URLs or local files or straight to cells of local Excel sheets.
#'
#' Note that for an external URL, only `file` and `text` should be supplied.
#' You can supply `dims` to `wb_add_formula()` to control the location of the link.
#' @param sheet Name of a worksheet
#' @param row integer row number for hyperlink to link to
#' @param col column number of letter for hyperlink to link to
#' @param text Display text
#' @param file Hyperlink or Excel file name to point to. If `NULL`, hyperlink is internal.
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
#'
#' ## Link to an URL.
#' x <- create_hyperlink(text = "openxlsx2 website", file = "https://janmarvin.github.io/openxlsx2/")
#'
#' wb$add_formula(x = x, dims = "A8")
#' # if (interactive()) wb$open()
#'
#' @seealso [wb_add_hyperlink()]
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

    if (is_waiver(sheet)) {
      sheet <- paste0('<<', toupper(sheet), '>>')
    }

    cell <- paste0(int2col(col), row)

    if (!is.null(file)) {
      dest <- sprintf('"[%s]%s!%s"', file, sheet, cell)
    } else {
      dest <- sprintf('"#\'%s\'!%s"', sheet, cell)
    }

    if (is.null(text)) {
      str <- sprintf("=HYPERLINK(%s)", dest)
    } else {
      str <- sprintf('=HYPERLINK(%s, \"%s\")', dest, text)
    }
  }

  return(str)
}

# `col2hex()` -----------------------------------------------------------------
#' Convert rgb to hex
#'
#' @param my.col my.col
#' @noRd
col2hex <- function(my.col) {
  grDevices::rgb(t(grDevices::col2rgb(my.col)), maxColorValue = 255)
}

# validate color --------------------------------------------------------------
#' Validate and convert color. Returns ARGB string
#' @param color input string (something like color(), "00000", "#000000", "00000000" or "#00000000")
#' @param or_null logical for use in assert
#' @param envir parent frame for use in assert
#' @param msg return message
#' @noRd
validate_color <- function(
  color = NULL, or_null = FALSE,
  envir = parent.frame(), msg = NULL,
  format = c("ARGB", "RGBA")
) {
  format <- match.arg(format)
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

  # if the format is RGBA (R's default), switch first two with last two characters
  if (format == "RGBA") {
    s <- nchar(color) == 8
    alpha <- substring(color[s], 7, 8)
    color[s] <- paste0(alpha, substring(color[s], 1, 6))
  }

  ## create a total size of 8 in ARGB format
  color <- stringi::stri_pad_left(str = color, width = 8, pad = "F")

  if (!all(grepl("[A-F0-9]{8}$", color))) {
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

  # Initialize an empty vector of three elements
  res <- c(L = "", C = "", R = "")

  has <- c(0, 0, 0)
  # Extract each component if present and remove &amp;[LCR]
  if (grepl("&amp;L", x)) {
    has[1] <- 1
    res[1] <- stringi::stri_extract_first_regex(x, "&amp;L[\\s\\S]*?(?=&amp;C|&amp;R|$)")
  }
  if (grepl("&amp;C", x)) {
    has[2] <- 1
    res[2] <- stringi::stri_extract_first_regex(x, "&amp;C[\\s\\S]*?(?=&amp;R|$)")
  }
  if (grepl("&amp;R", x)) {
    has[3] <- 1
    res[3] <- stringi::stri_extract_first_regex(x, "&amp;R[\\s\\S]*?$")
  }

  if (sum(has) == 0) return(character(0))

  res <- stringi::stri_replace_all_regex(res, "&amp;[LCR]", "")

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
#' @param scale_with_doc scale with doc
#' @param align_with_margins align with margins
#' @noRd
genHeaderFooterNode <- function(x, scale_with_doc = FALSE, align_with_margins = FALSE) {

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
    '<headerFooter differentOddEven="%s" differentFirst="%s" scaleWithDoc="%s" alignWithMargins="%s">',
    as.integer(!(is.null(evenHeader) & is.null(evenFooter))),
    as.integer(!(is.null(firstHeader) & is.null(firstFooter))),
    as_xml_attr(scale_with_doc),
    as_xml_attr(align_with_margins)
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

  xlsxFile
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

# Helper to split a cell range into rows or columns
split_dims <- function(dims, direction = "row") {
  df <- dims_to_dataframe(dims, fill = TRUE, empty_rm = TRUE)
  if (is.numeric(direction)) {
    if (direction == 1) direction <- "row"
    if (direction == 2) direction <- "col"
  }
  direction <- match.arg(direction, choices = c("row", "col"))
  if (direction == "row") df <- as.data.frame(t(df), stringsAsFactors = FALSE)
  vapply(df, FUN = function(x) {
    fst <- x[1]
    snd <- x[length(x)]
    sprintf("%s:%s", fst, snd)
  }, FUN.VALUE = NA_character_)
}

split_dim <- function(dims) {
  df <- dims_to_dataframe(dims, fill = TRUE, empty_rm = TRUE)
  if (ncol(df) > 1 && nrow(df) > 1)
    stop("`dims` should be a cell range of one row or one column.", call. = FALSE)
  unlist(df)
}

is_single_cell <- function(dims) {
  all(lengths(dims_to_rowcol(dims)) == 1)
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
#' @param type Either `NULL`, `stacked` or `column`
#' @param direction Either `NULL`, `row` (or `1`) or `col` (or `2`). Should
#' sparklines be created in the row or column direction? Defaults to `NULL`.
#' When `NULL` the direction is inferred from `dims` in cases where `dims`
#' spans a single row or column and defaults to `row` otherwise.
#' @param negative negative
#' @param display_empty_cells_as Either `gap`, `span` or `zero`
#' @param markers markers add marker to line
#' @param high highlight highest value
#' @param low highlight lowest value
#' @param first highlight first value
#' @param last highlight last value
#' @param color_series colorSeries
#' @param color_negative colorNegative
#' @param color_axis colorAxis
#' @param color_markers colorMarkers
#' @param color_first colorFirst
#' @param color_last colorLast
#' @param color_high colorHigh
#' @param color_low colorLow
#' @param manual_max manualMax
#' @param manual_min manualMin
#' @param line_weight lineWeight
#' @param date_axis dateAxis
#' @param display_x_axis displayXAxis
#' @param display_hidden displayHidden
#' @param min_axis_type minAxisType
#' @param max_axis_type maxAxisType
#' @param right_to_left rightToLeft
#' @param ... additional arguments
#' @return A string containing XML code
#' @examples
#' # create multiple sparklines
#' sparklines <- c(
#'   create_sparklines("Sheet 1", "A3:L3", "M3", type = "column", first = "1"),
#'   create_sparklines("Sheet 1", "A2:L2", "M2", markers = "1"),
#'   create_sparklines("Sheet 1", "A4:L4", "M4", type = "stacked", negative = "1"),
#'   create_sparklines("Sheet 1", "A5:L5;A7:L7", "M5;M7", markers = "1")
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
#' # create sparkline groups
#' sparklines <- c(
#'   create_sparklines("Sheet 2", "A2:L6;", "M2:M6", markers = "1"),
#'   create_sparklines(
#'     "Sheet 2", "A7:L7;A9:L9", "M7;M9", type = "stacked", negative = "1"
#'   ),
#'   create_sparklines(
#'     "Sheet 2", "A8:L8;A10:L13", "M8;M10:M13",
#'     type = "column", first = "1"
#'    ),
#'   create_sparklines(
#'     "Sheet 2", "A2:L13", "A14:L14", type = "column", first = "1",
#'     direction = "col"
#'   )
#' )
#'
#' wb <- wb$
#'   add_worksheet("Sheet 2")$
#'   add_data(x = t2)$
#'   add_sparklines(sparklines = sparklines)
#'
#' @export
create_sparklines <- function(
    sheet                  = current_sheet(),
    dims,
    sqref,
    type                   = NULL,
    negative               = NULL,
    display_empty_cells_as = "gap", # "span", "zero"
    markers                = NULL,
    high                   = NULL,
    low                    = NULL,
    first                  = NULL,
    last                   = NULL,
    color_series           = wb_color(hex = "FF376092"),
    color_negative         = wb_color(hex = "FFD00000"),
    color_axis             = wb_color(hex = "FFD00000"),
    color_markers          = wb_color(hex = "FFD00000"),
    color_first            = wb_color(hex = "FFD00000"),
    color_last             = wb_color(hex = "FFD00000"),
    color_high             = wb_color(hex = "FFD00000"),
    color_low              = wb_color(hex = "FFD00000"),
    manual_max             = NULL,
    manual_min             = NULL,
    line_weight            = NULL,
    date_axis              = NULL,
    display_x_axis         = NULL,
    display_hidden         = NULL,
    min_axis_type          = NULL,
    max_axis_type          = NULL,
    right_to_left          = NULL,
    direction              = NULL,
    ...
) {

  standardize_case_names(...)

  if (is_waiver(sheet)) {
    sheet <- paste0('<<', toupper(sheet), '>>')
  } else {
    assert_class(sheet, "character")
  }

  assert_class(dims,  "character")
  assert_class(sqref, "character")

  if (!is.null(type) && !type %in% c("stacked", "column"))
    stop("type must be NULL, stacked or column")

  if (!is.null(markers) && as_xml_attr(markers) == "" && !is.null(type) && type %in% c("stacked", "column"))
    stop("markers only affect lines `type = NULL`, not stacked or column")

  if (!is.null(direction) || !is_single_cell(sqref)) {
    dims <- split_dims(dims, direction = direction)
    sqref <- split_dim(sqref)
  }

  if (length(dims) != 1 && length(dims) != length(sqref)) {
    stop("dims and sqref must be equal length.")
  }

  sparklines <- vapply(
    seq_along(dims),
    function(i) {
      xml_node_create(
        "x14:sparkline",
        xml_children = c(
          xml_node_create(
            "xm:f",
            xml_children = c(
              paste0(shQuote(sheet, type = "sh"), "!", dims[[i]])
            )
          ),
          xml_node_create(
            "xm:sqref",
            xml_children = c(
              sqref[[i]]
            )
          )
        )
      )
    },
    FUN.VALUE = NA_character_
  )
  sparklines <- paste(sparklines, collapse = "")

  sparklineGroup <- xml_node_create(
    "x14:sparklineGroup",
    xml_attributes = c(
      type                = type,
      displayEmptyCellsAs = as_xml_attr(display_empty_cells_as),
      markers             = as_xml_attr(markers),
      high                = as_xml_attr(high),
      low                 = as_xml_attr(low),
      first               = as_xml_attr(first),
      last                = as_xml_attr(last),
      negative            = as_xml_attr(negative),
      manualMin           = as_xml_attr(manual_min),
      manualMax           = as_xml_attr(manual_max),
      lineWeight          = as_xml_attr(line_weight),
      dateAxis            = as_xml_attr(date_axis),
      displayXAxis        = as_xml_attr(display_x_axis),
      displayHidden       = as_xml_attr(display_hidden),
      minAxisType         = as_xml_attr(min_axis_type),
      maxAxisType         = as_xml_attr(max_axis_type),
      rightToLeft         = as_xml_attr(right_to_left),
      "xr2:uid"           = sprintf("{6F57B887-24F1-C14A-942C-%s}", random_string(length = 12, pattern = "[A-F0-9]"))
    ),
    xml_children = c(
      xml_node_create("x14:colorSeries",   xml_attributes = color_series),
      xml_node_create("x14:colorNegative", xml_attributes = color_negative),
      xml_node_create("x14:colorAxis",     xml_attributes = color_axis),
      xml_node_create("x14:colorMarkers",  xml_attributes = color_markers),
      xml_node_create("x14:colorFirst",    xml_attributes = color_first),
      xml_node_create("x14:colorLast",     xml_attributes = color_last),
      xml_node_create("x14:colorHigh",     xml_attributes = color_high),
      xml_node_create("x14:colorLow",      xml_attributes = color_low),
      xml_node_create(
        "x14:sparklines",
        xml_children = sparklines
      )
    )
  )

  sparklineGroup
}

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

#' Convert objects with attribute labels into strings. This allows special
#' type of labelled vectors, that assign labels to Inf/-Inf/NaN.
#' @param x a labelled vector to convert
#' @returns a character vector
#' @noRd
to_string <- function(x) {
  lbls <- attr(x, "labels")
  x_chr <- as.character(x)
  x_num <- suppressWarnings(as.numeric(x_chr))

  has_labels <- !is.null(lbls) && !is.null(names(lbls))
  used_label <- logical(length(x))

  if (has_labels) {
    idx <- match(x, lbls)
    has_label <- !is.na(idx)
    used_label <- has_label
    x_chr[has_label] <- names(lbls)[idx[has_label]]
  }

  # It is possible to assign labels to Inf, -Inf, and NaN. We have to
  # distinguish between a numeric Inf/-Inf/NaN and a label. The label
  # should be written as character, otherwise as #NUM! or #VALUE!
  if (any(!used_label)) {
    ul <- which(!used_label)

    inf_pos <- is.infinite(x_num[ul]) & x_num[ul] > 0
    inf_neg <- is.infinite(x_num[ul]) & x_num[ul] < 0
    is_nan  <- is.nan(x_num[ul])

    x_chr[ul[inf_pos]] <- "_openxlsx_Inf"
    x_chr[ul[inf_neg]] <- "_openxlsx_nInf"
    x_chr[ul[is_nan]]  <- "_openxlsx_NaN"
  }

  x_chr
}

# get the next free relationship id
get_next_id <- function(x, increase = 1L) {
  if (length(x)) {
    rlshp <- rbindlist(xml_attr(x, "Relationship"))
    rlshp$id <- as.integer(gsub("\\D+", "", rlshp$Id))
    next_id <- paste0("rId", max(rlshp$id) + increase)
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
    out
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
    out
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
    gsub(".*[\\/]", "", path)
  } else {
    basename(path)
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
    st_ids <- cols$style[cols$style != ""]
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

  sheet_id <- wb_old$clone()$.__enclos_env__$private$get_sheet_index(old)
  cc <- wb_old$worksheets[[sheet_id]]$sheet_data$cc
  sst_ids  <- as.integer(cc$v[cc$c_t == "s"]) + 1
  sst_uni  <- sort(unique(sst_ids))
  sst_old <- wb_old$sharedStrings[sst_uni]

  old_len <- length(as.character(wb_new$sharedStrings))

  wb_new$sharedStrings <- c(as.character(wb_new$sharedStrings), sst_old)
  attr(wb_new$sharedStrings, "uniqueCount") <- as.character(length(wb_new$sharedStrings))


  sheet_id <- wb_new$clone()$.__enclos_env__$private$get_sheet_index(new)
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

# In table names special characters must have a leading ' in front
# @params str a variable name
escape_specials <- function(str) {
  gsub("([^a-zA-Z0-9\\s])", "'\\1", str, perl = TRUE)
}

known_subtotal_funs <- function(x, total, table, row_names = FALSE) {

  # unfortunately x has no row names at this point
  ncol_x <- ncol(x) + row_names
  nms_x <- escape_specials(names(x))
  if (row_names) nms_x <- c("_rowNames_", nms_x)

  fml <- vector("character", ncol_x)
  atr <- vector("character", ncol_x)
  lbl <- rep(NA_character_, ncol_x)

  if (isTRUE(total) || all(as.character(total) == "109") || all(total == "sum")) {
    fml <- paste0("SUBTOTAL(109,", table, "[", nms_x, "])")
    atr <- rep("sum", ncol_x)
  } else {

    # all get the same total_row value
    if (length(total) == 1) {
      total <- rep(total, ncol_x)
    }

    if (length(total) != ncol_x) {
      stop("length of total_row and table columns do not match", call. = FALSE)
    }

    builtinIds <- c("101", "103", "102", "104", "105", "107", "109", "110")
    builtins   <- c("average", "count", "countNums", "max", "min", "stdDev", "sum", "var")

    ttl <- as.character(total)

    for (i in seq_len(ncol_x)) {

      if (any(names(total)[i] == "") && (ttl[i] %in% builtinIds || ttl[i] %in% builtins)) {
        if (ttl[i] == "101" || ttl[i] == "average") {
          fml[i] <- paste0("SUBTOTAL(", 101, ",", table, "[", nms_x[i], "])")
          atr[i] <- "average"
        } else if (ttl[i] == "102" || ttl[i] == "countNums") {
          fml[i] <- paste0("SUBTOTAL(", 102, ",", table, "[", nms_x[i], "])")
          atr[i] <- "countNums"
        } else if (ttl[i] == "103" || ttl[i] == "count") {
          fml[i] <- paste0("SUBTOTAL(", 103, ",", table, "[", nms_x[i], "])")
          atr[i] <- "count"
        } else if (ttl[i] == "104" || ttl[i] == "max") {
          fml[i] <- paste0("SUBTOTAL(", 104, ",", table, "[", nms_x[i], "])")
          atr[i] <- "max"
        } else if (ttl[i] == "105" || ttl[i] == "min") {
          fml[i] <- paste0("SUBTOTAL(", 105, ",", table, "[", nms_x[i], "])")
          atr[i] <- "min"
        } else if (ttl[i] == "107" || ttl[i] == "stdDev") {
          fml[i] <- paste0("SUBTOTAL(", 107, ",", table, "[", nms_x[i], "])")
          atr[i] <- "stdDev"
        } else if (ttl[i] == "109" || ttl[i] == "sum") {
          fml[i] <- paste0("SUBTOTAL(", 109, ",", table, "[", nms_x[i], "])")
          atr[i] <- "sum"
        } else if (ttl[i] == "110" || ttl[i] == "var") {
          fml[i] <- paste0("SUBTOTAL(", 110, ",", table, "[", nms_x[i], "])")
          atr[i] <- "var"
        }

      } else if (ttl[i] == "0" || ttl[i] == "none") {
        fml[i] <- ""
        atr[i] <- "none"
      } else if (any(names(total)[i] == "text")) {
        fml[i] <- as_xml_attr(ttl[i])
        atr[i] <- ""
        lbl[i] <- as_xml_attr(ttl[i])
      } else {
        # works, but in excel the formula is added to tables.xml as a child to the column
        fml[i] <- paste0(ttl[i], "(", table, "[", nms_x[i], "])")
        atr[i] <- "custom"
      }

    }

  }

  # prepare output
  fml <- as.data.frame(t(fml), stringsAsFactors = FALSE)
  names(fml) <- nms_x
  names(atr) <- nms_x
  names(lbl) <- nms_x

  # prepare output to be written with formulas
  for (i in seq_along(fml)) {
    if (is.na(lbl[[i]])) class(fml[[i]]) <- c("formula", fml[[i]])
  }

  list(fml, atr, lbl)

}

#' helper to read sheetPr xml to dataframe
#' @param xml xml_node
#' @noRd
read_sheetpr <- function(xml) {
  # https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetproperties?view=openxml-2.8.1
  if (!inherits(xml, "pugi_xml")) xml <- read_xml(xml)

  vec_attrs <- c("codeName", "enableFormatConditionsCalculation", "filterMode",
                 "published", "syncHorizontal", "syncRef", "syncVertical",
                 "transitionEntry", "transitionEvaluation")
  vec_chlds <- c("tabColor", "outlinePr", "pageSetUpPr")

  read_xml2df(
    xml       = xml,
    vec_name  = "sheetPr",
    vec_attrs = vec_attrs,
    vec_chlds = vec_chlds
  )
}

#' helper to write sheetPr dataframe to xml
#' @param df a data frame
#' @noRd
write_sheetpr <- function(df) {

  # we have to preserve a certain order of elements at least for childs
  vec_attrs <- c("codeName", "enableFormatConditionsCalculation", "filterMode",
                 "published", "syncHorizontal", "syncRef", "syncVertical",
                 "transitionEntry", "transitionEvaluation")
  vec_chlds <- c("tabColor", "outlinePr", "pageSetUpPr")
  nms <- c(vec_attrs, vec_chlds)

  write_df2xml(
    df        = df[nms],
    vec_name  = "sheetPr",
    vec_attrs = vec_attrs,
    vec_chlds = vec_chlds
  )
}

# helper construct dim comparison from rowcol_to_dims(as_integer = TRUE) object
min_and_max <- function(x) {
  c(
    (max(x[[2]]) + 1L) - min(x[[2]]), # row
    (max(x[[1]]) + 1L) - min(x[[1]])  # col
  )
}

#' helper function to detect if x fits into dims
#'
#' This function will throw a warning depending on the experimental option: `openxlsx2.warn_if_dims_dont_fit`
#' @param x the x object
#' @param dims the worksheet dimensions
#' @param startCol,startRow start column. Since write_data() is not defunct, we might not be fully able to select this from dims
#' @noRd
fits_in_dims <- function(x, dims, startCol, startRow) {

  # happens only in direct calls to write_data2 in some old tests
  if (is.null(dims)) {
    dims <- wb_dims(from_col = startCol, from_row = startRow)
  }

  if (length(dims) == 1 && is.character(dims)) {
    dims <- dims_to_rowcol(dims, as_integer = TRUE)
  }

  dim_x <- dim(x)
  dim_d <- min_and_max(dims)

  opt <- getOption("openxlsx2.warn_if_dims_dont_fit", default = FALSE)

  if (all(dim_x <= dim_d)) {
    fits <- TRUE
  } else if (all(dim_x > dim_d)) {
    if (opt) warning("dimension of `x` exceeds all `dims`")
    fits <- FALSE
  } else if (dim_x[1] > dim_d[1]) {
    if (opt) warning("dimension of `x` exceeds rows of `dims`")
    fits <- FALSE
  } else if (dim_x[2] > dim_d[2]) {
    if (opt) warning("dimension of `x` exceeds cols of `dims`")
    fits <- FALSE
  }

  if (fits) {

    # why oh why wasn't dims_to_rowcol()/rowcol_to_dims() created as a matching pair
    dims <- rowcol_to_dims(row = dims[["row"]], col = dims[["col"]])

  } else {

    # # one off. needs check if dims = NULL or row names argument?
    # dims <- wb_dims(x = x, from_col = startCol, from_row = startRow)

    data_nrow <- NROW(x)
    data_ncol <- NCOL(x)

    endRow <- (startRow - 1) + data_nrow
    endCol <- (startCol - 1) + data_ncol

    dims <- paste0(
      int2col(startCol), startRow,
      ":",
      int2col(endCol), endRow
    )

  }

  rc <- dims_to_rowcol(dims)
  if (max(as.integer(rc[["row"]])) > 1048576 || max(col2int(rc[["col"]])) > 16384)
    stop("Dimensions exceed worksheet")

  dims
}

# transpose single column or row data frames to wide/long. keeps attributes and
# class.
# The magic of t(). A Date can be something like a numeric with a
# format attached. After t(x) it will be a string "yyyy-mm-dd".
# Therefore unclass first and apply the class afterwards.
transpose_df <- function(x) {
  attribs <- attr(x, "c_cm")
  classes <- class(x[[1]])
  x[] <- lapply(x[], unclass)
  x <- as.data.frame(t(x), stringsAsFactors = FALSE)
  for (i in seq_along(x)) {
    class(x[[i]]) <- classes
  }
  attr(x, "c_cm") <- attribs
  x
}

#' helper function to update custom pids. Pids are indexed starting with 2.
#' @param wb a workbook
#' @noRd
wb_upd_custom_pid <- function(wb) {

  cstm <- xml_node(wb$custom, "Properties", "property")

  cstm_nams <- xml_node_name(cstm, "property")

  cstm_df      <- rbindlist(xml_attr(cstm, "property"))
  cstm_df$clds <- vapply(seq_along(cstm), function(x) xml_node(cstm[x], "property", cstm_nams[x]), NA_character_)
  cstm_df$pid  <- as.character(2L + (seq_len(nrow(cstm_df)) - 1L))

  out <- NULL
  for (i in seq_len(nrow(cstm_df))) {
    tmp <- xml_node_create(
      "property",
      xml_attributes = c(fmtid = cstm_df$fmtid[i], pid = cstm_df$pid[i], name = cstm_df$name[i]),
      xml_children   = c(cstm_df$clds[i])
    )
    out <- c(out, tmp)
  }

  ## return
  xml_node_create(
    "Properties",
    xml_attributes = c(
      xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
      `xmlns:vt` = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    ),
    xml_children = out
  )
}

#' replace shared formulas with single cell formulas
#' @param cc the full frame
#' @param cc_shared a subset of the full frame with shared formulas
#' @noRd
shared_as_fml <- function(cc, cc_shared) {

  ff <- rbindlist(xml_attr(paste0("<f ", cc_shared$f_attr, "/>"), "f"))
  cc_shared$f_si <- ff$si

  cc_shared <- cc_shared[order(as.integer(cc_shared$f_si)), ]

  # carry forward the shared formula
  cc_shared$f    <- ave2(cc_shared$f, cc_shared$f_si, carry_forward)

  # calculate differences from the formula cell, to the shared cells
  cc_shared$cols <- ave2(col2int(cc_shared$c_r), cc_shared$f_si, calc_distance)
  cc_shared$rows <- ave2(as.integer(cc_shared$row_r), cc_shared$f_si, calc_distance)

  # begin updating the formulas. find a1 notion, get the next cell, update formula
  cells <- find_a1_notation(cc_shared$f)
  repls <- vector("list", length = length(cells))

  for (i in seq_along(cells)) {
    repls[[i]] <- next_cell(cells[[i]], cc_shared$cols[i], cc_shared$rows[i])
  }

  cc_shared$f      <- replace_a1_notation(cc_shared$f, repls)
  cc_shared$cols   <- NULL
  cc_shared$rows   <- NULL
  cc_shared$f_attr <- rep("", nrow(cc_shared))
  cc_shared$f_si   <- NULL

  # reduce and assign
  cc_shared <- cc_shared[which(cc_shared$r %in% cc$r), ]

  cc[match(cc_shared$r, cc$r), names(cc_shared)] <- cc_shared
  cc
}

#' create a color used in create_shape
#' @param color a [wb_color()] object
#' @param transparency an integer value
#' @noRd
get_color <- function(color, transparency = 0) {

  alignment_map <- c(
    "0" =   "bg1",
    "1" =   "tx1",
    "2" =   "bg2",
    "3" =   "tx2",
    "4" =   "accent1",
    "5" =   "accent2",
    "6" =   "accent3",
    "7" =   "accent4",
    "8" =   "accent5",
    "9" =   "accent6",
    "10" =  "hlink",
    "11" =  "folHlink",
    "12" =  "phClr",
    "13" =  "dk1",
    "14" =  "lt1",
    "15" =  "dk2",
    "16" =  "lt2"
  )

  if (is_wbColour(color)) {
    if ("rgb" %in% names(color)) {
      color <- sprintf(
        '<a:solidFill>
        <a:srgbClr val="%s">
          <a:alpha val="%s" />
        </a:srgbClr>
      </a:solidFill>',
        substr(c(color["rgb"]), 3, 8),
        min(99, (100 - transparency)) * 1000
      )
    } else if ("theme" %in% names(color)) {
      color <- sprintf(
        '<a:solidFill>
        <a:schemeClr val="%s">
          <a:alpha val="%s" />
        </a:schemeClr>
      </a:solidFill>',
        alignment_map[color["theme"]],
        min(99, (100 - transparency)) * 1000
      )
    } else {
      warning("currently only rgb and theme colors are supported")
      color <- ""
    }
  } else {
    color <- ""
  }
  color
}

#' string styling used in create_shape()
#'
#' handles bold, italic, strike, size, font, charset
#' unhandled charset, outline, vert_align
#' @param txt input, character or [fmt_txt()]
#' @param text_color a [wb_color()]
#' @param transparency an integer value
#' @noRd
fmt_txt2 <- function(txt, text_color = "", transparency = 0) {
  if (!inherits(txt, "fmt_txt")) {
    txt <- fmt_txt(txt)
  }

  txts <- xml_node(txt, "r")

  out <- NULL
  for (txt in txts) { # no need to check for <b val="1"/>
    bold      <- ifelse(grepl("<b/>", txt), "1", "")
    italic    <- ifelse(grepl("<i/>", txt), "1", "")
    strike    <- ifelse(grepl("<strike/>", txt), "sngStrike", "")
    underline <- ifelse(grepl("<u/>", txt), "sng", "")

    color     <- sapply(xml_attr(txt, "r", "rPr", "color"), "[")
    if (length(color) == 0) {
      color   <- get_color(text_color, transparency)
    } else {
      color     <- get_color(wb_color(color), transparency) # tint?
    }

    sz <- sapply(xml_attr(txt, "r", "rPr", "sz"), "[")
    if (length(sz)) sz        <- as.integer(sz[["val"]]) * 100

    font <- sapply(xml_attr(txt, "r", "rPr", "rFont"), "[")
    charset <- sapply(xml_attr(txt, "r", "rPr", "charset"), "[")

    if (length(charset) == 0) charset <- c(val = "0")
    if (length(font)) {
      font <- c(
        sprintf('<a:latin typeface="%s" charset="%s" />', font[["val"]], charset[["val"]]),
        sprintf('<a:cs typeface="%s" charset="%s" />', font[["val"]], charset[["val"]])
      )
    } else {
      font <- NULL
    }

    rPr <- xml_node_create(
      "a:rPr",
      xml_attributes = c(
        b = as_xml_attr(bold),
        i = as_xml_attr(italic),
        sz = as_xml_attr(sz),
        strike = as_xml_attr(strike),
        u  = as_xml_attr(underline)
      ),
      xml_children = c(color, font)
    )


    text <- xml_value(txt, "r", "t")
    text <- xml_node_create("a:t", xml_children = text)
    ar   <- xml_node_create("a:r", xml_children = c(rPr, text))

    out <- c(out, ar)
  }

  paste0(out, collapse = "")
}

#' Helper to create a shape
#' @param shape a shape (see details)
#' @param name a name for the shape
#' @param text a text written into the object. This can be a simple character or a [fmt_txt()]
#' @param fill_color,text_color,line_color a color for each, accepts only theme and rgb colors passed with [wb_color()]
#' @param fill_transparency,text_transparency,line_transparency sets the alpha value of the shape, an integer value in the range 0 to 100
#' @param text_align sets the alignment of the text. Can be 'left', 'center', 'right', 'justify', 'justifyLow', 'distributed', or 'thaiDistributed'
#' @param rotation the rotation of the shape in degrees
#' @param id an integer id (effect is unknown)
#' @param ... additional arguments
#' @returns a character containing the XML
#' @seealso [wb_add_drawing()]
#' @details Possible shapes are (from ST_ShapeType - Preset Shape Types):
#' "line", "lineInv", "triangle", "rtTriangle", "rect", "diamond",
#' "parallelogram", "trapezoid", "nonIsoscelesTrapezoid", "pentagon",
#' "hexagon", "heptagon", "octagon", "decagon", "dodecagon", "star4",
#' "star5", "star6", "star7", "star8", "star10", "star12", "star16",
#' "star24", "star32", "roundRect", "round1Rect", "round2SameRect",
#' "round2DiagRect", "snipRoundRect", "snip1Rect", "snip2SameRect",
#' "snip2DiagRect", "plaque", "ellipse", "teardrop", "homePlate",
#' "chevron", "pieWedge", "pie", "blockArc", "donut", "noSmoking",
#' "rightArrow", "leftArrow", "upArrow", "downArrow", "stripedRightArrow",
#' "notchedRightArrow", "bentUpArrow", "leftRightArrow", "upDownArrow",
#' "leftUpArrow", "leftRightUpArrow", "quadArrow", "leftArrowCallout",
#' "rightArrowCallout", "upArrowCallout", "downArrowCallout", "leftRightArrowCallout",
#' "upDownArrowCallout", "quadArrowCallout", "bentArrow", "uturnArrow",
#' "circularArrow", "leftCircularArrow", "leftRightCircularArrow",
#' "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow",
#' "swooshArrow", "cube", "can", "lightningBolt", "heart", "sun",
#' "moon", "smileyFace", "irregularSeal1", "irregularSeal2", "foldedCorner",
#' "bevel", "frame", "halfFrame", "corner", "diagStripe", "chord",
#' "arc", "leftBracket", "rightBracket", "leftBrace", "rightBrace",
#' "bracketPair", "bracePair", "straightConnector1", "bentConnector2",
#' "bentConnector3", "bentConnector4", "bentConnector5", "curvedConnector2",
#' "curvedConnector3", "curvedConnector4", "curvedConnector5", "callout1",
#' "callout2", "callout3", "accentCallout1", "accentCallout2", "accentCallout3",
#' "borderCallout1", "borderCallout2", "borderCallout3", "accentBorderCallout1",
#' "accentBorderCallout2", "accentBorderCallout3", "wedgeRectCallout",
#' "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout",
#' "cloud", "ribbon", "ribbon2", "ellipseRibbon", "ellipseRibbon2",
#' "leftRightRibbon", "verticalScroll", "horizontalScroll", "wave",
#' "doubleWave", "plus", "flowChartProcess", "flowChartDecision",
#' "flowChartInputOutput", "flowChartPredefinedProcess", "flowChartInternalStorage",
#' "flowChartDocument", "flowChartMultidocument", "flowChartTerminator",
#' "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation",
#' "flowChartConnector", "flowChartPunchedCard", "flowChartPunchedTape",
#' "flowChartSummingJunction", "flowChartOr", "flowChartCollate",
#' "flowChartSort", "flowChartExtract", "flowChartMerge", "flowChartOfflineStorage",
#' "flowChartOnlineStorage", "flowChartMagneticTape", "flowChartMagneticDisk",
#' "flowChartMagneticDrum", "flowChartDisplay", "flowChartDelay",
#' "flowChartAlternateProcess", "flowChartOffpageConnector", "actionButtonBlank",
#' "actionButtonHome", "actionButtonHelp", "actionButtonInformation",
#' "actionButtonForwardNext", "actionButtonBackPrevious", "actionButtonEnd",
#' "actionButtonBeginning", "actionButtonReturn", "actionButtonDocument",
#' "actionButtonSound", "actionButtonMovie", "gear6", "gear9", "funnel",
#' "mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual",
#' "mathNotEqual", "cornerTabs", "squareTabs", "plaqueTabs", "chartX",
#' "chartStar", "chartPlus"
#'
#' @examples
#'  wb <- wb_workbook()$add_worksheet()$
#'    add_drawing(xml = create_shape())
#' @export
create_shape <- function(
    shape = "rect",
    name = "shape 1",
    text = "",
    fill_color = NULL,
    fill_transparency = 0,
    text_color = NULL,
    text_transparency = 0,
    line_color = fill_color,
    line_transparency = 0,
    text_align = "left",
    rotation = 0,
    id = 1,
    ...
) {

  valid_align <- c("left", "center", "right", "justify", "justifyLow", "distributed",
                   "thaiDistributed")
  match.arg(text_align, valid_align)

  valid_shapes <- c(
    "line", "lineInv", "triangle", "rtTriangle", "rect", "diamond",
    "parallelogram", "trapezoid", "nonIsoscelesTrapezoid", "pentagon",
    "hexagon", "heptagon", "octagon", "decagon", "dodecagon", "star4",
    "star5", "star6", "star7", "star8", "star10", "star12", "star16",
    "star24", "star32", "roundRect", "round1Rect", "round2SameRect",
    "round2DiagRect", "snipRoundRect", "snip1Rect", "snip2SameRect",
    "snip2DiagRect", "plaque", "ellipse", "teardrop", "homePlate",
    "chevron", "pieWedge", "pie", "blockArc", "donut", "noSmoking",
    "rightArrow", "leftArrow", "upArrow", "downArrow", "stripedRightArrow",
    "notchedRightArrow", "bentUpArrow", "leftRightArrow", "upDownArrow",
    "leftUpArrow", "leftRightUpArrow", "quadArrow", "leftArrowCallout",
    "rightArrowCallout", "upArrowCallout", "downArrowCallout", "leftRightArrowCallout",
    "upDownArrowCallout", "quadArrowCallout", "bentArrow", "uturnArrow",
    "circularArrow", "leftCircularArrow", "leftRightCircularArrow",
    "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow",
    "swooshArrow", "cube", "can", "lightningBolt", "heart", "sun",
    "moon", "smileyFace", "irregularSeal1", "irregularSeal2", "foldedCorner",
    "bevel", "frame", "halfFrame", "corner", "diagStripe", "chord",
    "arc", "leftBracket", "rightBracket", "leftBrace", "rightBrace",
    "bracketPair", "bracePair", "straightConnector1", "bentConnector2",
    "bentConnector3", "bentConnector4", "bentConnector5", "curvedConnector2",
    "curvedConnector3", "curvedConnector4", "curvedConnector5", "callout1",
    "callout2", "callout3", "accentCallout1", "accentCallout2", "accentCallout3",
    "borderCallout1", "borderCallout2", "borderCallout3", "accentBorderCallout1",
    "accentBorderCallout2", "accentBorderCallout3", "wedgeRectCallout",
    "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout",
    "cloud", "ribbon", "ribbon2", "ellipseRibbon", "ellipseRibbon2",
    "leftRightRibbon", "verticalScroll", "horizontalScroll", "wave",
    "doubleWave", "plus", "flowChartProcess", "flowChartDecision",
    "flowChartInputOutput", "flowChartPredefinedProcess", "flowChartInternalStorage",
    "flowChartDocument", "flowChartMultidocument", "flowChartTerminator",
    "flowChartPreparation", "flowChartManualInput", "flowChartManualOperation",
    "flowChartConnector", "flowChartPunchedCard", "flowChartPunchedTape",
    "flowChartSummingJunction", "flowChartOr", "flowChartCollate",
    "flowChartSort", "flowChartExtract", "flowChartMerge", "flowChartOfflineStorage",
    "flowChartOnlineStorage", "flowChartMagneticTape", "flowChartMagneticDisk",
    "flowChartMagneticDrum", "flowChartDisplay", "flowChartDelay",
    "flowChartAlternateProcess", "flowChartOffpageConnector", "actionButtonBlank",
    "actionButtonHome", "actionButtonHelp", "actionButtonInformation",
    "actionButtonForwardNext", "actionButtonBackPrevious", "actionButtonEnd",
    "actionButtonBeginning", "actionButtonReturn", "actionButtonDocument",
    "actionButtonSound", "actionButtonMovie", "gear6", "gear9", "funnel",
    "mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual",
    "mathNotEqual", "cornerTabs", "squareTabs", "plaqueTabs", "chartX",
    "chartStar", "chartPlus"
  )
  match.arg(shape, valid_shapes)

  alignment_map <- c(
    "left" = "l",
    "center" = "ctr",
    "right" = "r",
    "justify" = "just",
    "justifyLow" = "justLow",
    "distributed" = "dist",
    "thaiDistributed" = "thaiDist"
  )
  text_align <- alignment_map[text_align]

  standardize(...)

  if (!is_xml(text) || inherits(text, "fmt_txt")) {
    # if not a a14:m node
    text <- fmt_txt2(text, text_color = text_color, text_transparency)
    mc_beg <- ""
    mc_end <- ""
  } else {
    # we need some markup compability
    mc_beg <- c(
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
    )
    mc_end <- c(
      '</mc:Choice>
      </mc:AlternateContent>'
    )
  }

  line_color <- get_color(line_color, line_transparency)
  fill_color <- get_color(fill_color, fill_transparency)

  if (line_color != "") {
    line_color <- sprintf('<a:ln>%s</a:ln>', line_color)
  }

  xml <- sprintf(
    '<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">
     <xdr:absoluteAnchor>
      <xdr:pos x="0" y="0" />
      <xdr:ext cx="0" cy="0" />
      %s
      <xdr:sp macro="" textlink="">
       <xdr:nvSpPr>
        <xdr:cNvPr id="%s" name="%s" />
         <a:extLst>
          <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
           <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="%s" />
          </a:ext>
         </a:extLst>
        <xdr:cNvSpPr />
       </xdr:nvSpPr>
       <xdr:spPr>
        <a:xfrm rot="%s">
         <a:off x="0" y="0" />
         <a:ext cx="0" cy="0" />
        </a:xfrm>
        <a:prstGeom prst="%s">
         <a:avLst />
        </a:prstGeom>
        %s
        %s
       </xdr:spPr>
       <xdr:style>
        <a:lnRef idx="2">
         <a:schemeClr val="accent1">
          <a:shade val="50000" />
         </a:schemeClr>
        </a:lnRef>
        <a:fillRef idx="1">
         <a:schemeClr val="accent1" />
        </a:fillRef>
        <a:effectRef idx="0">
         <a:schemeClr val="accent1" />
        </a:effectRef>
        <a:fontRef idx="minor">
         <a:schemeClr val="lt1" />
        </a:fontRef>
       </xdr:style>
       <xdr:txBody>
        <a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t" />
        <a:lstStyle />
        <a:p>
         <a:pPr algn="%s" />
         %s
        </a:p>
       </xdr:txBody>
      </xdr:sp>
      %s
      <xdr:clientData />
     </xdr:absoluteAnchor>
     </xdr:wsDr>',
     mc_beg,
     id, name, st_guid(), rotation * 60000, shape,
     fill_color, line_color, text_align[1], text,
     mc_end
  )

  read_xml(xml, pointer = FALSE)
}
