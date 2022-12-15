wb_add_named_region_impl = function(
    self,
    private,
    sheet = current_sheet(),
    cols,
    rows,
    name,
    localSheet        = FALSE,
    overwrite         = FALSE,
    comment           = NULL,
    customMenu        = NULL,
    description       = NULL,
    is_function       = NULL,
    functionGroupId   = NULL,
    help              = NULL,
    hidden            = NULL,
    localName         = NULL,
    publishToServer   = NULL,
    statusBar         = NULL,
    vbProcedure       = NULL,
    workbookParameter = NULL,
    xml               = NULL
) {

  sheet <- private$get_sheet_index(sheet)

  if (!is.numeric(rows)) {
    stop("rows argument must be a numeric/integer vector")
  }

  if (!is.numeric(cols)) {
    stop("cols argument must be a numeric/integer vector")
  }

  localSheetId <- ""
  if (localSheet) localSheetId <- as.character(sheet)

  ## check name doesn't already exist
  ## named region

  definedNames <- rbindlist(xml_attr(self$workbook$definedNames, level1 = "definedName"))
  sel1 <- tolower(definedNames$name) == tolower(name)
  sel2 <- definedNames$localSheetId == localSheetId
  if (!is.null(definedNames$localSheetId)) {
    sel <- sel1 & sel2
  } else {
    sel <- sel1
  }
  match_dn <- which(sel)

  if (any(match_dn)) {
    if (overwrite)
      self$workbook$definedNames <- self$workbook$definedNames[-match_dn]
    else
      stop(sprintf("Named region with name '%s' already exists! Use overwrite  = TRUE if you want to replace it", name))
  } else if (grepl("^[A-Z]{1,3}[0-9]+$", name)) {
    stop("name cannot look like a cell reference.")
  }

  cols <- round(cols)
  rows <- round(rows)

  startCol <- min(cols)
  endCol <- max(cols)

  startRow <- min(rows)
  endRow <- max(rows)

  ref1 <- paste0("$", int2col(startCol), "$", startRow)
  ref2 <- paste0("$", int2col(endCol), "$", endRow)

  if (localSheetId == "") localSheetId <- NULL

  private$create_named_region(
    ref1               = ref1,
    ref2               = ref2,
    name               = name,
    sheet              = self$sheet_names[sheet],
    localSheetId       = localSheetId,
    comment            = comment,
    customMenu         = customMenu,
    description        = description,
    is_function        = is_function,
    functionGroupId    = functionGroupId,
    help               = help,
    hidden             = hidden,
    localName          = localName,
    publishToServer    = publishToServer,
    statusBar          = statusBar,
    vbProcedure        = vbProcedure,
    workbookParameter  = workbookParameter,
    xml                = xml
  )

  invisible(self)
}

wb_remove_named_region_impl = function(
    self,
    private,
    sheet = current_sheet(),
    name = NULL
) {
  # get all nown defined names
  dn <- get_named_regions(self)

  if (is.null(name) && !is.null(sheet)) {
    sheet <- private$get_sheet_index(sheet)
    del <- dn$id[dn$sheet == sheet]
  } else if (!is.null(name) && is.null(sheet)) {
    del <- dn$id[dn$name == name]
  } else {
    sheet <- private$get_sheet_index(sheet)
    del <- dn$id[dn$sheet == sheet & dn$name == name]
  }

  if (length(del)) {
    self$workbook$definedNames <- self$workbook$definedNames[-del]
  } else {
    if (!is.null(name))
      warning(sprintf("Cannot find named region with name '%s'", name))
    # do not warn if wb and sheet are selected. wb_delete_named_region is
    # called with every wb_remove_worksheet and would throw meaningless
    # warnings. For now simply assume if no name is defined, that the
    # user does not care, as long as no defined name remains on a sheet.
  }

  invisible(self)
}
