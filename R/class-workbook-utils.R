#' Validate sheet
#'
#' @param wb A workbook
#' @param sheetName The sheet name to validate
#' @return The sheet name -- or the position?  This should be consistent
#' @noRd
wb_validate_sheet <- function(wb, sheetName) {
  assert_workbook(wb)

  if (!is.numeric(sheetName)) {
    if (is.null(wb$sheet_names)) {
      stop("wb does not contain any worksheets.", call. = FALSE)
    }
  }

  if (is.numeric(sheetName)) {
    if (sheetName > length(wb$sheet_names)) {
      msg <- sprintf("wb only contains %i sheets.", length(wb$sheet_names))
      stop(msg, call. = FALSE)
    }

    return(sheetName)
  }

  if (!sheetName %in% replaceXMLEntities(wb$sheet_names)) {
    msg <- sprintf("Sheet '%s' does not exist.", replaceXMLEntities(sheetName))
    stop(msg, call. = FALSE)
  }


  which(replaceXMLEntities(wb$sheet_names) == sheetName)
}


#' Create a font node from a style
#'
#' @param wb a workbook
#' @param style style
#' @return The font node as xml?
#' @noRd
wb_create_font_node <- function(wb, style) {
  assert_workbook(wb)
  # assert_style(style)

  baseFont <- wb$getBaseFont()

  # TODO implement paste_c()
  fontNode <- "<font>"

  ## size
  if (is.null(style$fontSize[[1]])) {
    fontNode <- stri_join(fontNode, sprintf('<sz %s="%s"/>', names(baseFont$size), baseFont$size))
  } else {
    fontNode <- stri_join(fontNode, sprintf('<sz %s="%s"/>', names(style$fontSize), style$fontSize))
  }

  ## colour
  if (is.null(style$fontColour[[1]])) {
    fontNode <- stri_join(
      fontNode,
      sprintf('<color %s="%s"/>', names(baseFont$colour), baseFont$colour)
    )
  } else {
    if (length(style$fontColour) > 1) {
      fontNode <- stri_join(
        fontNode,
        sprintf(
          "<color %s/>",
          stri_join(
            sapply(
              seq_along(style$fontColour),
              function(i) {
                sprintf('%s="%s"', names(style$fontColour)[i], style$fontColour[i])
              }
            ),
            sep = " ",
            collapse = " "
          )
        )
      )
    } else {
      fontNode <- stri_join(
        fontNode,
        sprintf('<color %s="%s"/>', names(style$fontColour), style$fontColour)
      )
    }
  }


  ## name
  if (is.null(style$fontName[[1]])) {
    fontNode <- stri_join(
      fontNode,
      sprintf('<name %s="%s"/>', names(baseFont$name), baseFont$name)
    )
  } else {
    fontNode <- stri_join(
      fontNode,
      sprintf('<name %s="%s"/>', names(style$fontName), style$fontName)
    )
  }

  ### Create new font and return Id
  if (!is.null(style$fontFamily)) {
    fontNode <- stri_join(fontNode, sprintf('<family val = "%s"/>', style$fontFamily))
  }

  if (!is.null(style$fontScheme)) {
    fontNode <- stri_join(fontNode, sprintf('<scheme val = "%s"/>', style$fontScheme))
  }

  if ("BOLD" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, "<b/>")
  }

  if ("ITALIC" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, "<i/>")
  }

  if ("UNDERLINE" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, '<u val="single"/>')
  }

  if ("UNDERLINE2" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, '<u val="double"/>')
  }

  if ("STRIKEOUT" %in% style$fontDecoration) {
    fontNode <- stri_join(fontNode, "<strike/>")
  }

  stri_join(fontNode, "</font>")
}

#' Validate table name for a workbook
#'
#' @param wb a workbook
#' @param tableName Table name
#' @return A valid table name as a `character`
#' @noRd
wb_validate_table_name <- function(wb, tableName) {
  assert_workbook(wb)
  # returns the new tableName -- basically just lowercase
  tableName <- tolower(tableName) ## Excel forces named regions to lowercase

  # TODO set these to warnings? trim and peplace bad characters with

  # TODO add a strict = getOption("openxlsx.tableName.strict", FALSE)
  # param to force these to allow to stopping
  if (nchar(tableName) > 255) {
    stop("tableName must be less than 255 characters.", call. = FALSE)
  }

  if (grepl("$", tableName, fixed = TRUE)) {
    stop("'$' character cannot exist in a tableName", call. = FALSE)
  }

  if (grepl(" ", tableName, fixed = TRUE)) {
    stop("spaces cannot exist in a table name", call. = FALSE)
  }

  # if(!grepl("^[A-Za-z_]", tableName, perl = TRUE))
  #   stop("tableName must begin with a letter or an underscore")

  if (grepl("R[0-9]+C[0-9]+", tableName, perl = TRUE, ignore.case = TRUE)) {
    stop("tableName cannot be the same as a cell reference, such as R1C1", call. = FALSE)
  }

  if (grepl("^[A-Z]{1,3}[0-9]+$", tableName, ignore.case = TRUE)) {
    stop("tableName cannot be the same as a cell reference", call. = FALSE)
  }

  # only place where self is needed
  if (tableName %in% attr(wb$tables, "tableName")) {
    stop(sprintf("table with name '%s' already exists", tableName), call. = FALSE)
  }

  tableName
}

#' Checks for overwrite columns
#'
#' @param wb workbook
#' @param sheet sheet
#' @param new_rows new_rows
#' @param new_cols new_cols
#' @param error_msg error_msg
#' @param check_table_header_only check_table_header_only
#' @noRd
wb_check_overwrite_tables <- function(
  wb,
  sheet,
  new_rows,
  new_cols,
  # why does error_msg need to be a param?
  error_msg = "Cannot overwrite existing table with another table.",
  check_table_header_only = FALSE
) {
  # TODO pull out -- no assignemnts made
  ## check not overwriting another table
  if (length(wb$tables)) {
    tableSheets <- attr(wb$tables, "sheet")
    sheetNo <- wb_validate_sheet(wb, sheet)

    to_check <-
      which(tableSheets %in% sheetNo &
          !grepl("openxlsx_deleted", attr(wb$tables, "tableName"), fixed = TRUE))

    if (length(to_check)) {
      ## only look at tables on this sheet

      exTable <- wb$tables[to_check]

      rows <- lapply(
        names(exTable),
        function(rectCoords) {
          as.numeric(unlist(regmatches(rectCoords, gregexpr("[0-9]+", rectCoords))))
        }
      )
      cols <- lapply(
        names(exTable),
        function(rectCoords) {
          col2int(unlist(regmatches(rectCoords, gregexpr("[A-Z]+", rectCoords))))
        }
      )

      if (check_table_header_only) {
        rows <- lapply(rows, function(x) c(x[1], x[1]))
      }


      ## loop through existing tables checking if any over lap with new table
      for (i in seq_along(exTable)) {
        existing_cols <- cols[[i]]
        existing_rows <- rows[[i]]

        if ((min(new_cols) <= max(existing_cols)) &
            (max(new_cols) >= min(existing_cols)) &
            (min(new_rows) <= max(existing_rows)) &
            (max(new_rows) >= min(existing_rows))) {
          stop(error_msg)
        }
      }
    } ## end if(sheet %in% tableSheets)
  } ## end (length(tables))

  invisible(wb)
}
