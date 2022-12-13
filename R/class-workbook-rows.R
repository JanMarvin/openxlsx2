wb_set_row_heights_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    rows,
    heights
) {
  sheet <- private$get_sheet_index(sheet)

  # TODO move to wbWorksheet method

  # create all A columns so that row_attr is available
  dims <- rowcol_to_dims(rows, 1)
  private$do_cell_init(sheet, dims)

  if (length(rows) > length(heights)) {
    heights <- rep(heights, length.out = length(rows))
  }

  if (length(heights) > length(rows)) {
    stop("Greater number of height values than rows.")
  }

  row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

  sel <- match(rows, row_attr$r)
  row_attr[sel, "ht"] <- as.character(as.numeric(heights))
  row_attr[sel, "customHeight"] <- "1"

  self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

  invisible(self)
}

wb_remove_row_heights_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    rows
  ) {
  sheet <- private$get_sheet_index(sheet)

  row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

  if (is.null(row_attr)) {
    warning("There are no initialized rows on this sheet")
    return(invisible(self))
  }

  sel <- match(rows, row_attr$r)
  row_attr[sel, "ht"] <- ""
  row_attr[sel, "customHeight"] <- ""

  self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

  invisible(self)
}

wb_group_rows_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    rows,
    collapsed = FALSE,
    levels = NULL
) {

  sheet <- private$get_sheet_index(sheet)

  if (length(collapsed) > length(rows)) {
    stop("Collapses argument is of greater length than number of rows.")
  }

  if (!is.logical(collapsed)) {
    stop("Collapses should be a logical value (TRUE/FALSE).")
  }

  if (any(rows <= 0L)) {
    stop("Invalid rows entered (<= 0).")
  }

  collapsed <- rep(as.character(as.integer(collapsed)), length.out = length(rows))

  levels <- levels %||% rep("1", length(rows))

  # Remove duplicates
  ok <- !duplicated(rows)
  collapsed <- collapsed[ok]
  levels <- levels[ok]
  rows <- rows[ok]
  sheet <- private$get_sheet_index(sheet)

  # fetch the row_attr data.frame
  row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

  rows_rev <- rev(rows)

  # get the selection based on the row_attr frame.

  # the first n -1 rows get outlineLevel
  select <- row_attr$r %in% as.character(rows_rev[-1])
  if (length(select)) {
    row_attr$outlineLevel[select] <- as.character(levels[-1])
    row_attr$collapsed[select] <- as.character(as.integer(collapsed[-1]))
    row_attr$hidden[select] <- as.character(as.integer(collapsed[-1]))
  }

  # the n-th row gets only collapsed
  select <- row_attr$r %in% as.character(rows_rev[1])
  if (length(select)) {
    row_attr$collapsed[select] <- as.character(as.integer(collapsed[1]))
  }

  self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

  # check if there are valid outlineLevel in row_attr and assign outlineLevelRow the max outlineLevel (thats in the documentation)
  if (any(row_attr$outlineLevel != "")) {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = as.character(max(as.integer(row_attr$outlineLevel), na.rm = TRUE))))
  }

  invisible(self)
}

wb_ungroup_rows_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    rows
) {
  sheet <- private$get_sheet_index(sheet)

  # check if any rows are selected
  if (any(rows < 1L)) {
    stop("Invalid rows entered (<= 0).")
  }

  # fetch the row_attr data.frame
  row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

  # get the selection based on the row_attr frame.
  select <- row_attr$r %in% as.character(rows)
  if (length(select)) {
    row_attr$outlineLevel[select] <- ""
    row_attr$collapsed[select] <- ""
    # TODO only if unhide = TRUE
    row_attr$hidden[select] <- ""
    self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr
  }

  # If all outlineLevels are missing: remove the outlineLevelRow attribute. Assigning "" will remove the attribute
  if (all(row_attr$outlineLevel == "")) {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = ""))
  } else {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = as.character(max(as.integer(row_attr$outlineLevel)))))
  }

  invisible(self)
}
