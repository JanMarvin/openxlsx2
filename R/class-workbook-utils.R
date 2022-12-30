#' Validate sheet
#'
#' @param wb A workbook
#' @param sheet The sheet name to validate
#' @return The sheet name -- or the position?  This should be consistent
#' @noRd
wb_validate_sheet <- function(wb, sheet) {
  assert_workbook(wb)
  wb$validate_sheet(sheet)
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

  # TODO add a strict = getOption("openxlsx2.tableName.strict", FALSE)
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

  # if (!grepl("^[A-Za-z_]", tableName, perl = TRUE))
  #   stop("tableName must begin with a letter or an underscore")

  if (grepl("R[0-9]+C[0-9]+", tableName, perl = TRUE, ignore.case = TRUE)) {
    stop("tableName cannot be the same as a cell reference, such as R1C1", call. = FALSE)
  }

  if (grepl("^[A-Z]{1,3}[0-9]+$", tableName, ignore.case = TRUE)) {
    stop("tableName cannot be the same as a cell reference", call. = FALSE)
  }

  # only place where self is needed
  if (tableName %in% wb$tables$tab_name) {
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
  if (!is.null(wb$tables)) {
    tableSheets <- wb$tables$tab_sheet
    sheetNo <- wb_validate_sheet(wb, sheet)

    to_check <- tableSheets %in% sheetNo & wb$tables$tab_act == 1

    if (length(to_check)) {
      ## only look at tables on this sheet

      exTable <- wb$tables[to_check, ]

      exTable$rows <- lapply(
        exTable$tab_ref,
        function(rectCoords) {
          as.numeric(unlist(regmatches(rectCoords, gregexpr("[0-9]+", rectCoords))))
        }
      )
      exTable$cols <- lapply(
        exTable$tab_ref,
        function(rectCoords) {
          col2int(unlist(regmatches(rectCoords, gregexpr("[A-Z]+", rectCoords))))
        }
      )

      if (check_table_header_only) {
        exTable$rows <- lapply(exTable$rows, function(x) c(x[1], x[1]))
      }


      ## loop through existing tables checking if any over lap with new table
      for (i in seq_len(NROW(exTable))) {
        existing_cols <- exTable$cols[[i]]
        existing_rows <- exTable$rows[[i]]

        if ((min(new_cols) <= max(existing_cols)) &&
            (max(new_cols) >= min(existing_cols)) &&
            (min(new_rows) <= max(existing_rows)) &&
            (max(new_rows) >= min(existing_rows))) {
          stop(error_msg)
        }
      }
    } ## end if (sheet %in% tableSheets)
  } ## end (length(tables))

  invisible(wb)
}


validate_cf_params <- function(params) {
  bad <- names(params) %out% c("border", "gradient", "iconSet", "percent", "rank", "reverse", "showValue")
  if (any(bad)) {
    stop("Invalid parameters: ", toString(names(params)[bad]))
  }

  # assign default values
  params$showValue <- if (is.null(params$showValue)) 1L else as_binary(params$showValue)
  params$gradient  <- if (is.null(params$gradient))  1L else as_binary(params$gradient)
  params$border    <- if (is.null(params$border))    1L else as_binary(params$border)
  params$percent   <- if (is.null(params$percent))   0L else as_binary(params$percent)

  # special check for rank
  params$rank <- params$rank %||% 5L

  if (!is_integer_ish(params$rank)) {
    stop("params$rank must be an integer")
  }

  params$rank <- as.integer(params$rank)

  params
}

# waivers -----------------------------------------------------------------

current_sheet <- function() {
  structure("current_sheet", class = "openxlsx2_waiver")
}

next_sheet <- function() {
  structure("next_sheet", class = "openxlsx2_waiver")
}

is_waiver <- function(x) {
  inherits(x, "openxlsx2_waiver")
}
