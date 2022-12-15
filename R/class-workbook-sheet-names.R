wb_get_sheet_names_impl <- function(self, private) {
  res <- self$sheet_names
  names(res) <- private$original_sheet_names
  res[self$sheetOrder]
}

wb_set_sheet_names_impl <- function(self, private, old = NULL, new) {
  # assume all names.  Default values makes the test check for wrappers a
  # little weird
  old <- old %||% seq_along(self$sheet_names)

  if (identical(old, new)) {
    return(self)
  }

  if (!length(self$worksheets)) {
    stop("workbook does not contain any sheets")
  }

  if (length(old) != length(new)) {
    stop("`old` and `new` must be the same length")
  }

  pos <- private$get_sheet_index(old)
  new_raw <- as.character(new)
  new_name <- replace_legal_chars(new_raw)

  if (identical(self$sheet_names[pos], new_name)) {
    return(self)
  }

  bad <- duplicated(tolower(new))
  if (any(bad)) {
    stop("Sheet names cannot have duplicates: ", toString(new[bad]))
  }

  # should be able to pull this out into a single private function
  for (i in seq_along(pos)) {
    private$validate_new_sheet(new_name[i])
    private$set_single_sheet_name(pos[i], new_name[i], new_raw[i])
    # TODO move this work into private$set_single_sheet_name()

    ## Rename in workbook
    sheetId <- private$get_sheet_id(type = "sheetId", pos[i])
    rId <- private$get_sheet_id(type = 'rId', pos[i])
    self$workbook$sheets[[pos[i]]] <-
      sprintf(
        '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
        new_name[i],
        sheetId,
        rId
      )

    ## rename defined names
    if (length(self$workbook$definedNames)) {
      ind <- get_named_regions(self)$sheets == old
      if (any(ind)) {
        nn <- sprintf("'%s'", new_name[i])
        nn <- stringi::stri_replace_all_fixed(self$workbook$definedName[ind], old, nn)
        nn <- stringi::stri_replace_all(nn, "'+", "'")
        self$workbook$definedNames[ind] <- nn
      }
    }
  }

  invisible(self)
}
