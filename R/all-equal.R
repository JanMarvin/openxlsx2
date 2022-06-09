#' @method all.equal wbWorkbook
#' @export
all.equal.wbWorkbook <- function(target, current, ...) {


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

  assert_workbook(target)
  assert_workbook(current)

  x <- target
  y <- current

  along <- seq_along((x))
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


  flag <- sapply(along, function(i) !isTRUE(all.equal(x$worksheets[[i]]$sheet_data$cc, y$worksheets[[i]]$sheet_data$cc)))
  if (any(flag)) {
    for (i in which(flag[flag])) {
      txt <- sprintf("cc elements on sheet %s: not equal",i)
      message(txt)
      failures <- c(failures, txt)
    }
  }


  flag <- sapply(along, function(i) !isTRUE(all.equal(x$worksheets[[i]]$sheet_data$row_attr, y$worksheets[[i]]$sheet_data$row_attr)))
  if (any(flag)) {
    for (i in which(flag[flag])) {
      txt <- sprintf("row_attr elements on sheet %s: not equal",i)
      message(txt)
      failures <- c(failures, txt)
    }
  }

  flag <- all(x$styles_mgr$styles %in% y$styles_mgr$styles) & all(y$styles_mgr$styles %in% x$styles_mgr$styles)
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

expect_equal_workbooks <- function(object, expected, ..., ignore_fields = NULL) {
  # Quick internal expectation function.  Easier to ignore fields to be set

  requireNamespace("testthat")
  requireNamespace("waldo")

  # object <- as.list(object)
  # expected <- as.list(expected)

  assert_workbook(object)
  assert_workbook(expected)

  fields <- c(names(wbWorkbook$public_fields), names(wbWorkbook$private_fields))
  bad <- setdiff(ignore_fields, fields)

  if (length(bad)) {
    stop("Invalid fields: ", toString(bad))
  }

  for (i in ignore_fields) {
    object[[i]] <- NULL
    expected[[i]] <- NULL
  }

  bad <- waldo::compare(
    x     = object,
    y     = expected,
    x_arg = "object",
    y_arg = "expected",
    ...
  )

  if (length(bad)) {
    testthat::fail(bad)
    return(invisible())
  }

  testthat::succeed()
  return(invisible())
}
