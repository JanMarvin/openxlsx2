class_workbook_add_data_validations <- function(
    self,
    private,
    sheet        = current_sheet(),
    cols,
    rows,
    type,
    operator,
    value,
    allowBlank   = TRUE,
    showInputMsg = TRUE,
    showErrorMsg = TRUE,
    errorStyle   = NULL,
    errorTitle   = NULL,
    error        = NULL,
    promptTitle  = NULL,
    prompt       = NULL
) {
  sheet <- private$get_sheet_index(sheet)

  ## rows and cols
  if (!is.numeric(cols)) {
    cols <- col2int(cols)
  }
  rows <- as.integer(rows)

  assert_class(allowBlank, "logical")
  assert_class(showInputMsg, "logical")
  assert_class(showErrorMsg, "logical")

  ## check length of value
  if (length(value) > 2) {
    stop("value argument must be length <= 2")
  }

  valid_types <- c(
    "custom",
    "whole",
    "decimal",
    "date",
    "time", ## need to conv
    "textLength",
    "list"
  )

  if (!tolower(type) %in% tolower(valid_types)) {
    stop("Invalid 'type' argument!")
  }

  ## operator == 'between' we leave out
  valid_operators <- c(
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual"
  )

  if (!tolower(type) %in% c("custom", "list")) {
    if (!tolower(operator) %in% tolower(valid_operators)) {
      stop("Invalid 'operator' argument!")
    }

    operator <- valid_operators[tolower(valid_operators) %in% tolower(operator)][1]
  } else if (tolower(type) == "custom") {
    operator <- NULL
  } else {
    operator <- "between" ## ignored
  }

  ## All inputs validated

  type <- valid_types[tolower(valid_types) %in% tolower(type)][1]

  ## check input combinations
  if ((type == "date") && !inherits(value, "Date")) {
    stop("If type == 'date' value argument must be a Date vector.")
  }

  if ((type == "time") && !inherits(value, c("POSIXct", "POSIXt"))) {
    stop("If type == 'time' value argument must be a POSIXct or POSIXlt vector.")
  }


  value <- head(value, 2)
  allowBlank <- as.character(as.integer(allowBlank[1]))
  showInputMsg <- as.character(as.integer(showInputMsg[1]))
  showErrorMsg <- as.character(as.integer(showErrorMsg[1]))

  # prepare for worksheet
  origin <- get_date_origin(self, origin = TRUE)

  sqref <- stri_join(
    get_cell_refs(data.frame(
      "x" = c(min(rows), max(rows)),
      "y" = c(min(cols), max(cols))
    )),
    sep = " ",
    collapse = ":"
  )

  if (type == "list") {
    operator <- NULL
  }

  self$worksheets[[sheet]]$.__enclos_env__$private$data_validation(
    type         = type,
    operator     = operator,
    value        = value,
    allowBlank   = allowBlank,
    showInputMsg = showInputMsg,
    showErrorMsg = showErrorMsg,
    errorStyle   = errorStyle,
    errorTitle   = errorTitle,
    error        = error,
    promptTitle  = promptTitle,
    prompt       = prompt,
    origin       = origin,
    sqref        = sqref
  )

  invisible(self)
}

workbook_merge_cells <- function(
    self,
    private,
    sheet = current_sheet(),
    rows  = NULL,
    cols  = NULL
) {
  sheet <- private$get_sheet_index(sheet)

  # TODO send to wbWorksheet() method
  # self$worksheets[[sheet]]$merge_cells(rows = rows, cols = cols)
  # invisible(self)

  rows <- range(as.integer(rows))
  cols <- range(as.integer(cols))

  sqref <- paste0(int2col(cols), rows)
  sqref <- stri_join(sqref, collapse = ":", sep = " ")

  current <- rbindlist(xml_attr(xml = self$worksheets[[sheet]]$mergeCells, "mergeCell"))$ref

  # regmatch0 will return character(0) when x is NULL
  if (length(current)) {

    new_merge     <- unname(unlist(dims_to_dataframe(sqref, fill = TRUE)))
    current_cells <- lapply(current, function(x) unname(unlist(dims_to_dataframe(x, fill = TRUE))))
    intersects    <- vapply(current_cells, function(x) any(x %in% new_merge), NA)

    # Error if merge intersects
    if (any(intersects)) {
      msg <- sprintf(
        "Merge intersects with existing merged cells: \n\t\t%s.\nRemove existing merge first.",
        stri_join(current[intersects], collapse = "\n\t\t")
      )
      stop(msg, call. = FALSE)
    }
  }

  # TODO does this have to be xml?  Can we just save the data.frame or
  # matrix and then check that?  This would also simplify removing the
  # merge specifications
  private$append_sheet_field(sheet, "mergeCells", sprintf('<mergeCell ref="%s"/>', sqref))
  invisible(self)
}

workbook_unmerge_cells <- function(
    self,
    private,
    sheet = current_sheet(),
    rows = NULL,
    cols = NULL
) {
  sheet <- private$get_sheet_index(sheet)

  rows <- range(as.integer(rows))
  cols <- range(as.integer(cols))

  sqref <- paste0(int2col(cols), rows)
  sqref <- stri_join(sqref, collapse = ":", sep = " ")

  current <- rbindlist(xml_attr(xml = self$worksheets[[sheet]]$mergeCells, "mergeCell"))$ref

  if (!is.null(current)) {
    new_merge     <- unname(unlist(dims_to_dataframe(sqref, fill = TRUE)))
    current_cells <- lapply(current, function(x) unname(unlist(dims_to_dataframe(x, fill = TRUE))))
    intersects    <- vapply(current_cells, function(x) any(x %in% new_merge), NA)

    # Remove intersection
    self$worksheets[[sheet]]$mergeCells <- self$worksheets[[sheet]]$mergeCells[!intersects]
  }

  invisible(self)
}

workbook_free_panes <- function(
    self,
    private,
    sheet          = current_sheet(),
    firstActiveRow = NULL,
    firstActiveCol = NULL,
    firstRow       = FALSE,
    firstCol       = FALSE
) {
  # TODO rename to setFreezePanes?

  # fine to do the validation before the actual check to prevent other errors
  sheet <- private$get_sheet_index(sheet)

  if (is.null(firstActiveRow) & is.null(firstActiveCol) & !firstRow & !firstCol) {
    return(invisible(self))
  }

  # TODO simplify asserts
  if (!is.logical(firstRow)) stop("firstRow must be TRUE/FALSE")
  if (!is.logical(firstCol)) stop("firstCol must be TRUE/FALSE")

  # make overwrides for arguments
  if (firstRow & !firstCol) {
    firstActiveCol <- NULL
    firstActiveRow <- NULL
    firstCol <- FALSE
  } else if (firstCol & !firstRow) {
    firstActiveRow <- NULL
    firstActiveCol <- NULL
    firstRow <- FALSE
  } else if (firstRow & firstCol) {
    firstActiveRow <- 2L
    firstActiveCol <- 2L
    firstRow <- FALSE
    firstCol <- FALSE
  } else {
    ## else both firstRow and firstCol are FALSE
    firstActiveRow <- firstActiveRow %||% 1L
    firstActiveCol <- firstActiveCol %||% 1L

    # Convert to numeric if column letter given
    # TODO is col2int() safe for non characters?
    firstActiveRow <- col2int(firstActiveRow)
    firstActiveCol <- col2int(firstActiveCol)
  }

  paneNode <-
    if (firstRow) {
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
    } else if (firstCol) {
      '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>'
    } else {
      if (firstActiveRow == 1 & firstActiveCol == 1) {
        ## nothing to do
        # return(NULL)
        return(invisible(self))
      }

      if (firstActiveRow > 1 & firstActiveCol == 1) {
        attrs <- sprintf('ySplit="%s"', firstActiveRow - 1L)
        activePane <- "bottomLeft"
      }

      if (firstActiveRow == 1 & firstActiveCol > 1) {
        attrs <- sprintf('xSplit="%s"', firstActiveCol - 1L)
        activePane <- "topRight"
      }

      if (firstActiveRow > 1 & firstActiveCol > 1) {
        attrs <- sprintf('ySplit="%s" xSplit="%s"',
                         firstActiveRow - 1L,
                         firstActiveCol - 1L
        )
        activePane <- "bottomRight"
      }

      sprintf(
        '<pane %s topLeftCell="%s" activePane="%s" state="frozen"/><selection pane="%s"/>',
        stri_join(attrs, collapse = " ", sep = " "),
        get_cell_refs(data.frame(firstActiveRow, firstActiveCol)),
        activePane,
        activePane
      )
    }

  self$worksheets[[sheet]]$freezePane <- paneNode
  invisible(self)
}
