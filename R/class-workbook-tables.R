wb_get_tables_impl <- function(self, private, sheet = current_sheet()) {
  if (length(sheet) != 1) {
    stop("sheet argument must be length 1")
  }

  if (is.null(self$tables)) {
    return(character())
  }

  sheet <- private$get_sheet_index(sheet)
  if (is.na(sheet)) stop("No such sheet in workbook")

  sel <- self$tables$tab_sheet == sheet & self$tables$tab_act == 1
  tables <- self$tables$tab_name[sel]
  refs <- self$tables$tab_ref[sel]

  if (length(tables)) {
    attr(tables, "refs") <- refs
  }

  tables
}

wb_remove_tables_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    table
) {
  if (length(table) != 1) {
    stop("table argument must be length 1")
  }

  ## delete table object and all data in it
  sheet <- private$get_sheet_index(sheet)

  if (!table %in% self$tables$tab_name) {
    stop(sprintf("table '%s' does not exist.", table), call. = FALSE)
  }

  ## delete table object (by flagging as deleted)
  inds <- self$tables$tab_sheet %in% sheet & self$tables$tab_name %in% table
  table_name_original <- self$tables$tab_name[inds]
  refs <- self$tables$tab_ref[inds]

  self$tables$tab_name[inds] <- paste0(table_name_original, "_openxlsx_deleted")
  self$tables$tab_ref[inds] <- ""
  self$tables$tab_sheet[inds] <- 0
  self$tables$tab_xml[inds] <- ""
  self$tables$tab_act[inds] <- 0

  ## delete reference from worksheet to table
  worksheet_table_names <- attr(self$worksheets[[sheet]]$tableParts, "tableName")
  to_remove <- which(worksheet_table_names == table_name_original)

  # (1) remove the rId from worksheet_rels
  rm_tab_rId <- rbindlist(xml_attr(self$worksheets[[sheet]]$tableParts[to_remove], "tablePart"))["r:id"]
  ws_rels <- self$worksheets_rels[[sheet]]
  is_rm_table <- grepl(rm_tab_rId, ws_rels)
  self$worksheets_rels[[sheet]] <- ws_rels[!is_rm_table]

  # (2) remove the rId from tableParts
  self$worksheets[[sheet]]$tableParts <- self$worksheets[[sheet]]$tableParts[-to_remove]
  attr(self$worksheets[[sheet]]$tableParts, "tableName") <- worksheet_table_names[-to_remove]


  ## Now delete data from the worksheet
  refs <- strsplit(refs, split = ":")[[1]]
  rows <- as.integer(gsub("[A-Z]", "", refs))
  rows <- seq(from = rows[1], to = rows[2], by = 1)

  cols <- col2int(refs)
  cols <- seq(from = cols[1], to = cols[2], by = 1)

  ## now delete data
  delete_data(wb = self, sheet = sheet, rows = rows, cols = cols)
  invisible(self)
}
