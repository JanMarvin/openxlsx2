wb_add_cols_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    n,
    beg,
    end
) {
  sheet <- private$get_sheet_index(sheet)
  self$worksheets[[sheet]]$cols_attr <- df_to_xml("col", empty_cols_attr(n, beg, end))
  invisible(self)
}

wb_group_cols_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    cols,
    collapsed = FALSE,
    levels = NULL
) {
  sheet <- private$get_sheet_index(sheet)

  if (length(collapsed) > length(cols)) {
    stop("Collapses argument is of greater length than number of cols.")
  }

  if (!is.logical(collapsed)) {
    stop("Collapses should be a logical value (TRUE/FALSE).")
  }

  if (any(cols < 1L)) {
    stop("Invalid rows entered (<= 0).")
  }

  collapsed <- rep(as.character(as.integer(collapsed)), length.out = length(cols))
  levels <- levels %||% rep("1", length(cols))

  # Remove duplicates
  ok <- !duplicated(cols)
  collapsed <- collapsed[ok]
  levels    <- levels[ok]
  cols      <- cols[ok]

  # fetch the row_attr data.frame
  col_attr <- self$worksheets[[sheet]]$unfold_cols()

  if (NROW(col_attr) == 0) {
    # TODO should this be a warning?  Or an error?
    message("worksheet has no columns. please create some with createCols")
  }

  # reverse to make it easier to get the fist
  cols_rev <- rev(cols)

  # get the selection based on the col_attr frame.

  # the first n -1 cols get outlineLevel
  select <- col_attr$min %in% as.character(cols_rev[-1])
  if (length(select)) {
    col_attr$outlineLevel[select] <- as.character(levels[-1])
    col_attr$collapsed[select] <- as.character(as.integer(collapsed[-1]))
    col_attr$hidden[select] <- as.character(as.integer(collapsed[-1]))
  }

  # the n-th row gets only collapsed
  select <- col_attr$min %in% as.character(cols_rev[1])
  if (length(select)) {
    col_attr$collapsed[select] <- as.character(as.integer(collapsed[1]))
  }

  self$worksheets[[sheet]]$fold_cols(col_attr)


  # check if there are valid outlineLevel in col_attr and assign outlineLevelRow the max outlineLevel (thats in the documentation)
  if (any(col_attr$outlineLevel != "")) {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = as.character(max(as.integer(col_attr$outlineLevel), na.rm = TRUE))))
  }

  invisible(self)
}

wb_ungroup_cols_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    cols
) {
  sheet <- private$get_sheet_index(sheet)

  # check if any rows are selected
  if (any(cols < 1L)) {
    stop("Invalid cols entered (<= 0).")
  }

  # fetch the cols_attr data.frame
  col_attr <- self$worksheets[[sheet]]$unfold_cols()

  # get the selection based on the col_attr frame.
  select <- col_attr$min %in% as.character(cols)

  if (length(select)) {
    col_attr$outlineLevel[select] <- ""
    col_attr$collapsed[select] <- ""
    # TODO only if unhide = TRUE
    col_attr$hidden[select] <- ""
    self$worksheets[[sheet]]$fold_cols(col_attr)
  }

  # If all outlineLevels are missing: remove the outlineLevelCol attribute. Assigning "" will remove the attribute
  if (all(col_attr$outlineLevel == "")) {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = ""))
  } else {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = as.character(max(as.integer(col_attr$outlineLevel)))))
  }

  invisible(self)
}

wb_remove_col_widths_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    cols
) {
  sheet <- private$get_sheet_index(sheet)

  if (!is.numeric(cols)) {
    cols <- col2int(cols)
  }

  customCols <- as.integer(names(self$colWidths[[sheet]]))
  removeInds <- which(customCols %in% cols)
  if (length(removeInds)) {
    remainingCols <- customCols[-removeInds]
    if (length(remainingCols) == 0) {
      self$colWidths[[sheet]] <- list()
    } else {
      rem_widths <- self$colWidths[[sheet]][-removeInds]
      names(rem_widths) <- as.character(remainingCols)
      self$colWidths[[sheet]] <- rem_widths
    }
  }

  invisible(self)
}

wb_set_col_widths_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    cols,
    widths = 8.43,
    hidden = FALSE
) {
  sheet <- private$get_sheet_index(sheet)

  # should do nothing if the cols' length is zero
  # TODO why would cols ever be 0?  Can we just signal this as an error?
  if (length(cols) == 0L) {
    return(invisible(self))
  }

  cols <- col2int(cols)

  if (length(widths) > length(cols)) {
    stop("More widths than columns supplied.")
  }

  if (length(hidden) > length(cols)) {
    stop("hidden argument is longer than cols.")
  }

  if (length(widths) < length(cols)) {
    widths <- rep(widths, length.out = length(cols))
  }

  if (length(hidden) < length(cols)) {
    hidden <- rep(hidden, length.out = length(cols))
  }

  # TODO add bestFit option?
  bestFit <- rep("1", length.out = length(cols))
  customWidth <- rep("1", length.out = length(cols))

  ## Remove duplicates
  ok <- !duplicated(cols)
  col_width <- widths[ok]
  hidden <- hidden[ok]
  cols <- cols[ok]

  col_df <- self$worksheets[[sheet]]$unfold_cols()
  base_font <- self$get_base_font()

  if (any(widths == "auto")) {
    df <- wb_to_df(self, sheet = sheet, cols = cols, colNames = FALSE)
    # TODO format(x) might not be the way it is formatted in the xlsx file.
    col_width <- vapply(df, function(x) max(nchar(format(x))), NA_real_)
  }

  # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column
  widths <- calc_col_width(base_font = base_font, col_width = col_width)

  # create empty cols
  if (NROW(col_df) == 0) {
    self$createCols(sheet, n = max(cols))
    col_df <- col_to_df(read_xml(self$worksheets[[sheet]]$cols_attr))
  }

  # found a few cols, but not all required cols. create the missing columns
  if (any(!cols %in% as.numeric(col_df$min))) {
    beg <- max(as.numeric(col_df$min)) + 1
    end <- max(cols)

    # new columns
    self$createCols(sheet, beg = beg, end = end)
    new_cols <- col_to_df(read_xml(self$worksheets[[sheet]]$cols_attr))

    # rbind only the missing columns. avoiding dups
    sel <- !new_cols$min %in% col_df$min
    col_df <- rbind(col_df, new_cols[sel, ])
    col_df <- col_df[order(as.numeric(col_df[, "min"])), ]
  }

  select <- as.numeric(col_df$min) %in% cols
  col_df$width[select] <- widths
  col_df$hidden[select] <- tolower(hidden)
  col_df$bestFit[select] <- bestFit
  col_df$customWidth[select] <- customWidth
  self$worksheets[[sheet]]$fold_cols(col_df)
  invisible(self)
}
