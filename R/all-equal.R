#' @name all.equal
#' @aliases all.equal.Workbook
#' @title Check equality of workbooks
#' @description Check equality of workbooks
#' @method all.equal Workbook
#' @param target A `Workbook` object
#' @param current A `Workbook` object
#' @param ... ignored
all.equal.Workbook <- function(target, current, ...) {


  # print("Comparing workbooks...")
  #   ".rels",
  #   "app",
  #   "charts",
  #   "colWidths",
  #   "Content_Types",
  #   "core",
  #   "drawings",
  #   "drawings_rels",
  #   "media",
  #   "rowHeights",
  #   "workbook",
  #   "workbook.xml.rels",
  #   "worksheets",
  #   "sheetOrder"
  #   "sharedStrings",
  #   "tables",
  #   "tables.xml.rels",
  #   "theme"


  ## TODO
  # sheet_data

  x <- target
  y <- current




  along <- seq_along(names(x))
  failures <- NULL

  flag <- all(names(x$charts) %in% names(y$charts)) & all(names(y$charts) %in% names(x$charts))
  if (!flag) {
    message("charts not equal")
    failures <- c(failures, "wb$charts")
  }

  flag <- all(sapply(along, function(i) isTRUE(all.equal(x$colWidths[[i]], y$colWidths[[i]]))))
  if (!flag) {
    message("colWidths not equal")
    failures <- c(failures, "wb$colWidths")
  }

  flag <- all(x$Content_Types %in% y$Content_Types) & all(y$Content_Types %in% x$Content_Types)
  if (!flag) {
    message("Content_Types not equal")
    failures <- c(failures, "wb$Content_Types")
  }

  flag <- all(unlist(x$core) == unlist(y$core))
  if (!flag) {
    message("core not equal")
    failures <- c(failures, "wb$core")
  }


  flag <- all(unlist(x$drawings) %in% unlist(y$drawings)) & all(unlist(y$drawings) %in% unlist(x$drawings))
  if (!flag) {
    message("drawings not equal")
    failures <- c(failures, "wb$drawings")
  }

  flag <- all(unlist(x$drawings_rels) %in% unlist(y$drawings_rels)) & all(unlist(y$drawings_rels) %in% unlist(x$drawings_rels))
  if (!flag) {
    message("drawings_rels not equal")
    failures <- c(failures, "wb$drawings_rels")
  }

  flag <- all(sapply(along, function(i) isTRUE(all.equal(x$drawings_rels[[i]], y$drawings_rels[[i]]))))
  if (!flag) {
    message("drawings_rels not equal")
    failures <- c(failures, "wb$drawings_rels")
  }




  flag <- all(names(x$media) %in% names(y$media) & names(y$media) %in% names(x$media))
  if (!flag) {
    message("media not equal")
    failures <- c(failures, "wb$media")
  }

  flag <- all(sapply(along, function(i) isTRUE(all.equal(x$rowHeights[[i]], y$rowHeights[[i]]))))
  if (!flag) {
    message("rowHeights not equal")
    failures <- c(failures, "wb$rowHeights")
  }

  flag <- all(sapply(along, function(i) isTRUE(all.equal(names(x$rowHeights[[i]]), names(y$rowHeights[[i]])))))
  if (!flag) {
    message("rowHeights not equal")
    failures <- c(failures, "wb$rowHeights")
  }

  flag <- all(x$sharedStrings %in% y$sharedStrings) & all(y$sharedStrings %in% x$sharedStrings) & (length(x$sharedStrings) == length(y$sharedStrings))
  if (!flag) {
    message("sharedStrings not equal")
    failures <- c(failures, "wb$sharedStrings")
  }



  # flag <- sapply(along, function(i) isTRUE(all.equal(x$worksheets[[i]]$sheet_data, y$worksheets[[i]]$sheet_data)))
  # if(!all(flag)){
  #
  #   tmp_x <- x$sheet_data[[which(!flag)[[1]]]]
  #   tmp_y <- y$sheet_data[[which(!flag)[[1]]]]
  #
  #   tmp_x_e <- sapply(tmp_x, "[[", "r")
  #   tmp_y_e <- sapply(tmp_y, "[[", "r")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("r elements: %s", paste(which(flag), collapse = ", ")))
  #     return(FALSE)
  #   }
  #
  #   tmp_x_e <- sapply(tmp_x, "[[", "t")
  #   tmp_y_e <- sapply(tmp_y, "[[", "t")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("t elements: %s", paste(which(isTRUE(flag)), collapse = ", ")))
  #     return(FALSE)
  #   }
  #
  #
  #   tmp_x_e <- sapply(tmp_x, "[[", "v")
  #   tmp_y_e <- sapply(tmp_y, "[[", "v")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("v elements: %s", paste(which(flag), collapse = ", ")))
  #     return(FALSE)
  #   }
  #
  #   tmp_x_e <- sapply(tmp_x, "[[", "f")
  #   tmp_y_e <- sapply(tmp_y, "[[", "f")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("f elements: %s", paste(which(flag), collapse = ", ")))
  #     return(FALSE)
  #   }
  # }


  flag <- all(names(x$styles) %in% names(y$styles)) & all(names(y$styles) %in% names(x$styles))
  if (!flag) {
    message("names styles not equal")
    failures <- c(failures, "names of styles not equal")
  }

  flag <- all(unlist(x$styles) %in% unlist(y$styles)) & all(unlist(y$styles) %in% unlist(x$styles))
  if (!flag) {
    message("styles not equal")
    failures <- c(failures, "styles not equal")
  }


  flag <- all(x$sheet_names %in% y$sheet_names) & all(y$sheet_names %in% x$sheet_names)
  if (!flag) {
    message("names workbook not equal")
    failures <- c(failures, "names workbook not equal")
  }

  flag <- all(unlist(x$workbook) %in% unlist(y$workbook)) & all(unlist(y$workbook) %in% unlist(x$workbook))
  if (!flag) {
    message("workbook not equal")
    failures <- c(failures, "wb$workbook")
  }

  flag <- all(unlist(x$workbook.xml.rels) %in% unlist(y$workbook.xml.rels)) & all(unlist(y$workbook.xml.rels) %in% unlist(x$workbook.xml.rels))
  if (!flag) {
    message("workbook.xml.rels not equal")
    failures <- c(failures, "wb$workbook.xml.rels")
  }


  flag <- all(unlist(x$sheetOrder) %in% unlist(y$sheetOrder)) & all(unlist(y$sheetOrder) %in% unlist(x$sheetOrder))
  if (!flag) {
    message("sheetOrder not equal")
    failures <- c(failures, "sheetOrder not equal")
  }


  flag <- length(x$tables) == length(y$tables)
  if (!flag) {
    message("length of tables not equal")
    failures <- c(failures, "length of tables not equal")
  }

  flag <- all(names(x$tables) == names(y$tables))
  if (!flag) {
    message("names of tables not equal")
    failures <- c(failures, "names of tables not equal")
  }

  flag <- all(unlist(x$tables) == unlist(y$tables))
  if (!flag) {
    message("tables not equal")
    failures <- c(failures, "tables not equal")
  }


  flag <- isTRUE(all.equal(x$tables.xml.rels, y$tables.xml.rels))
  if (!flag) {
    message("tables.xml.rels not equal")
    failures <- c(failures, "tables.xml.rels not equal")
  }

  flag <- x$theme == y$theme
  if (!flag) {
    message("theme not equal")
    failures <- c(failures, "theme not equal")
  }

  if (!is.null(failures)) {
    return(FALSE)
  }


  #   "connections",
  #   "externalLinks",
  #   "externalLinksRels",
  #   "headFoot",
  #   "pivotTables",
  #   "pivotTables.xml.rels",
  #   "pivotDefinitions",
  #   "pivotRecords",
  #   "pivotDefinitionsRels",
  #   "queryTables",
  #   "slicers",
  #   "slicerCaches",
  #   "vbaProject",


  return(TRUE)
}
