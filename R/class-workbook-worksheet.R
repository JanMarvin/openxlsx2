workbook_remove_worksheet <- function(self, private, sheet = current_sheet()) {

  # To delete a worksheet
  # Remove colwidths element
  # Remove drawing partname from Content_Types (drawing(sheet).xml)
  # Remove highest sheet from Content_Types
  # Remove drawings element
  # Remove drawings_rels element

  # Remove vml element
  # Remove vml_rels element

  # Remove rowHeights element
  # Remove last sheet element from workbook
  # Remove last sheet element from workbook.xml.rels
  # Remove element from worksheets
  # Remove element from worksheets_rels
  # Remove hyperlinks
  # Reduce calcChain i attributes & remove calcs on sheet
  # Remove sheet from sheetOrder
  # Remove queryTable references from workbook$definedNames to worksheet
  # remove tables

  # TODO can we allow multiple sheets?
  if (length(sheet) != 1) {
    stop("sheet must have length 1.")
  }

  sheet       <- private$get_sheet_index(sheet)
  sheet_names <- self$sheet_names
  nSheets     <- length(sheet_names)
  sheet_names <- sheet_names[[sheet]]

  ## definedNames
  if (length(self$workbook$definedNames)) {
    # wb_validate_sheet() makes sheet an integer
    # so we need to remove this before getting rid of the sheet names
    self$workbook$definedNames <- self$workbook$definedNames[
      !get_nr_from_definedName(self)$sheets %in% self$sheet_names[sheet]
    ]
  }

  self$remove_named_region(sheet)
  self$sheet_names <- self$sheet_names[-sheet]
  private$original_sheet_names <- private$original_sheet_names[-sheet]

  # if a sheet has no relationships, we xml_rels will not contain data
  if (!is.null(self$worksheets_rels[[sheet]])) {

    xml_rels <- rbindlist(
      xml_attr(self$worksheets_rels[[sheet]], "Relationship")
    )

    if (nrow(xml_rels) && ncol(xml_rels)) {
      xml_rels$type   <- basename(xml_rels$Type)
      xml_rels$target <- basename(xml_rels$Target)
      xml_rels$target[xml_rels$type == "hyperlink"] <- ""
      xml_rels$target_ind <- as.numeric(gsub("\\D+", "", xml_rels$target))
    }

    comment_id    <- xml_rels$target_ind[xml_rels$type == "comments"]
    # TODO not every sheet has a drawing. this originates from a time where
    # every sheet created got a drawing assigned.
    drawing_id    <- xml_rels$target_ind[xml_rels$type == "drawing"]
    pivotTable_id <- xml_rels$target_ind[xml_rels$type == "pivotTable"]
    table_id      <- xml_rels$target_ind[xml_rels$type == "table"]
    thrComment_id <- xml_rels$target_ind[xml_rels$type == "threadedComment"]
    vmlDrawing_id <- xml_rels$target_ind[xml_rels$type == "vmlDrawing"]

    # NULL the sheets
    if (length(comment_id))    self$comments[[comment_id]]          <- NULL
    if (length(drawing_id))    self$drawings[[drawing_id]]          <- ""
    if (length(drawing_id))    self$drawings_rels[[drawing_id]]     <- ""
    if (length(thrComment_id)) self$threadComments[[thrComment_id]] <- NULL
    if (length(vmlDrawing_id)) self$vml[[vmlDrawing_id]]            <- NULL
    if (length(vmlDrawing_id)) self$vml_rels[[vmlDrawing_id]]       <- NULL

    #### Modify Content_Types
    ## remove last drawings(sheet).xml from Content_Types
    drawing_name <- xml_rels$target[xml_rels$type == "drawing"]
    if (!is.null(drawing_name) && !identical(drawing_name, character()))
      self$Content_Types <- grep(drawing_name, self$Content_Types, invert = TRUE, value = TRUE)

  }

  self$isChartSheet <- self$isChartSheet[-sheet]

  ## remove highest sheet
  # (don't chagne this to a "grep(value = TRUE)" ... )
  self$Content_Types <- self$Content_Types[!grepl(sprintf("sheet%s.xml", nSheets), self$Content_Types)]

  # The names for the other drawings have changed
  de <- xml_node(read_xml(self$Content_Types), "Default")
  ct <- rbindlist(xml_attr(read_xml(self$Content_Types), "Override"))
  ct[grepl("drawing", ct$PartName), "PartName"] <- sprintf("/xl/drawings/drawing%i.xml", seq_along(self$drawings))
  ct <- df_to_xml("Override", ct[c("PartName", "ContentType")])
  self$Content_Types <- c(de, ct)


  ## sheetOrder
  toRemove <- which(self$sheetOrder == sheet)
  self$sheetOrder[self$sheetOrder > sheet] <- self$sheetOrder[self$sheetOrder > sheet] - 1L
  self$sheetOrder <- self$sheetOrder[-toRemove]

  # TODO regexpr should be replaced
  ## Need to remove reference from workbook.xml.rels to pivotCache
  removeRels <- grep("pivotTables", self$worksheets_rels[[sheet]], value = TRUE)
  if (length(removeRels)) {
    ## sheet rels links to a pivotTable file, the corresponding pivotTable_rels file links to the cacheDefn which is listing in workbook.xml.rels
    ## remove reference to this file from the workbook.xml.rels
    fileNo <- reg_match0(removeRels, "(?<=pivotTable)[0-9]+(?=\\.xml)")
    fileNo <- as.integer(unlist(fileNo))

    toRemove <- stri_join(
      sprintf("(pivotCacheDefinition%i\\.xml)", fileNo),
      sep = " ",
      collapse = "|"
    )

    toRemove <- stri_join(
      sprintf("(pivotCacheDefinition%i\\.xml)", grep(toRemove, self$pivotTables.xml.rels)),
      sep = " ",
      collapse = "|"
    )

    ## remove reference to file from workbook.xml.res
    self$workbook.xml.rels <- grep(toRemove, self$workbook.xml.rels, invert = TRUE, value = TRUE)
  }

  ## As above for slicers
  ## Need to remove reference from workbook.xml.rels to pivotCache
  # removeRels <- grepl("slicers", self$worksheets_rels[[sheet]])

  if (any(grepl("slicers", self$worksheets_rels[[sheet]]))) {
    # don't change to a grep(value = TRUE)
    self$workbook.xml.rels <- self$workbook.xml.rels[!grepl(sprintf("(slicerCache%s\\.xml)", sheet), self$workbook.xml.rels)]
  }

  ## wont't remove tables and then won't need to reassign table r:id's but will rename them!
  self$worksheets[[sheet]] <- NULL
  self$worksheets_rels[[sheet]] <- NULL

  sel <- self$tables$tab_sheet == sheet
  # tableName is a character Vector with an attached name Vector.
  if (any(sel)) {
    self$tables$tab_name[sel] <- paste0(self$tables$tab_name[sel], "_openxlsx_deleted")
    tab_sheet <- self$tables$tab_sheet
    tab_sheet[sel] <- 0
    tab_sheet[tab_sheet > sheet] <- tab_sheet[tab_sheet > sheet] - 1L
    self$tables$tab_sheet <- tab_sheet
    self$tables$tab_ref[sel] <- ""
    self$tables$tab_xml[sel] <- ""

    # deactivate sheet
    self$tables$tab_act[sel] <- 0
  }

  # ## drawing will always be the first relationship
  # if (nSheets > 1) {
  #   for (i in seq_len(nSheets - 1L)) {
  #     # did this get updated from length of 3 to 2?
  #     #self$worksheets_rels[[i]][1:2] <- genBaseSheetRels(i)
  #     rel <- rbindlist(xml_attr(self$worksheets_rels[[i]], "Relationship"))
  #     if (nrow(rel) && ncol(rel)) {
  #       if (any(basename(rel$Type) == "drawing")) {
  #         rel$Target[basename(rel$Type) == "drawing"] <- sprintf("../drawings/drawing%s.xml", i)
  #       }
  #       if (is.null(rel$TargetMode)) rel$TargetMode <- ""
  #       self$worksheets_rels[[i]] <- df_to_xml("Relationship", rel[c("Id", "Type", "Target", "TargetMode")])
  #     }
  #   }
  # } else {
  #   self$worksheets_rels <- list()
  # }

  ## remove sheet
  sn <- apply_reg_match0(self$workbook$sheets, pat = '(?<= name=")[^"]+')
  self$workbook$sheets <- self$workbook$sheets[!sn %in% sheet_names]

  ## Reset rIds
  if (nSheets > 1) {
    for (i in (sheet + 1L):nSheets) {
      self$workbook$sheets <- gsub(
        stri_join("rId", i),
        stri_join("rId", i - 1L),
        self$workbook$sheets,
        fixed = TRUE
      )
    }
  } else {
    self$workbook$sheets <- NULL
  }

  ## Can remove highest sheet
  # (don't use grepl(value = TRUE))
  self$workbook.xml.rels <- self$workbook.xml.rels[!grepl(sprintf("sheet%s.xml", nSheets), self$workbook.xml.rels)]

  invisible(self)
}

