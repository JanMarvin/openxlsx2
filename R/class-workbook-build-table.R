workbook_build_table <- function(
    self,
    private,
    colNames,
    ref,
    showColNames,
    tableStyle,
    tableName,
    withFilter, # TODO set default for withFilter?
    totalsRowCount    = 0,
    showFirstColumn   = 0,
    showLastColumn    = 0,
    showRowStripes    = 1,
    showColumnStripes = 0
) {
  ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
  last_table_id <- function() {
    z <- 0

    if (!all(unlist(self$worksheets_rels) == "")) {
      relship <- rbindlist(xml_attr(unlist(self$worksheets_rels), "Relationship"))
      # assign("relship", relship, globalenv())
      relship$typ <- basename(relship$Type)
      relship$tid <- as.numeric(gsub("\\D+", "", relship$Target))
      if (any(relship$typ == "table"))
        z <- max(relship$tid[relship$typ == "table"])
    }

    z
  }

  id <- as.character(last_table_id() + 1) # otherwise will start at 0 for table 1 length indicates the last known
  sheet <- wb_validate_sheet(self, sheet)
  # get the next highest rid
  rid <- 1
  if (!all(identical(self$worksheets_rels[[sheet]], character()))) {
    rid <- max(as.integer(sub("\\D+", "", rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))[["Id"]]))) + 1
  }

  if (is.null(self$tables)) {
    nms <- NULL
    tSheets <- NULL
    tNames <- NULL
    tActive <- NULL
  } else {
    nms <- self$tables$tab_ref
    tSheets <- self$tables$tab_sheet
    tNames <- self$tables$tab_name
    tActive <- self$tables$tab_act
  }


  ### autofilter
  autofilter <- if (withFilter) {
    xml_node_create(xml_name = "autoFilter", xml_attributes = c(ref = ref))
  }

  ### tableColumn
  tableColumn <- sapply(colNames, function(x) {
    id <- which(colNames %in% x)
    xml_node_create("tableColumn", xml_attributes = c(id = id, name = x))
  })

  tableColumns <- xml_node_create(
    xml_name = "tableColumns",
    xml_children = tableColumn,
    xml_attributes = c(count = as.character(length(colNames)))
  )


  ### tableStyleInfo
  tablestyle_attr <- c(
    name              = tableStyle,
    showFirstColumn   = as.integer(showFirstColumn),
    showLastColumn    = as.integer(showLastColumn),
    showRowStripes    = as.integer(showRowStripes),
    showColumnStripes = as.integer(showColumnStripes)
  )

  tableStyleXML <- xml_node_create(xml_name = "tableStyleInfo", xml_attributes = tablestyle_attr)


  ### full table
  table_attrs <- c(
    xmlns          = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    `xmlns:mc`     = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    id             = id,
    name           = tableName,
    displayName    = tableName,
    ref            = ref,
    totalsRowCount = totalsRowCount,
    totalsRowShown = "0"
    #headerRowDxfId="1"
  )

  tab_xml_new <- xml_node_create(
    xml_name = "table",
    xml_children = c(autofilter, tableColumns, tableStyleXML),
    xml_attributes = table_attrs
  )

  self$tables <- data.frame(
    tab_name = c(tNames, tableName),
    tab_sheet = c(tSheets, sheet),
    tab_ref = c(nms, ref),
    tab_xml = c(self$tables$tab_xml, tab_xml_new),
    tab_act = c(self$tables$tab_act, 1),
    stringsAsFactors = FALSE
  )

  self$worksheets[[sheet]]$tableParts <- c(
    self$worksheets[[sheet]]$tableParts,
    sprintf('<tablePart r:id="rId%s"/>', rid)
  )
  attr(self$worksheets[[sheet]]$tableParts, "tableName") <- c(
    tNames[tSheets == sheet & tActive == 1],
    tableName
  )

  ## create a table.xml.rels
  self$append("tables.xml.rels", "")

  ## update worksheets_rels
  self$worksheets_rels[[sheet]] <- c(
    self$worksheets_rels[[sheet]],
    sprintf(
      '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
      rid,
      id
    )
  )

  invisible(self)
}
