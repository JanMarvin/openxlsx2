.use_collapse <- function() {
  isTRUE(getOption("openxlsx2.use_collapse", TRUE)) && requireNamespace("collapse", quietly = TRUE)
}

f_unique <- function(x, ...) {
  if (.use_collapse() && !is.null(x)) {
    collapse::funique(x, ...)
  } else {
    unique(x, ...)
  }
}

f_match <- function(x, table, nomatch = NA_integer_, incomparables = NULL) {
  if (.use_collapse() && is.null(incomparables)) {
    collapse::fmatch(x, table, nomatch = nomatch)
  } else {
    match(x, table, nomatch = nomatch, incomparables = incomparables)
  }
}

f_range <- function(x, ..., na.rm = TRUE) {
  if (.use_collapse() && length(list(...)) == 0) {
    collapse::frange(x, na.rm = na.rm)
  } else {
    range(x, ..., na.rm = na.rm)
  }
}

f_duplicated <- function(x, incomparables = FALSE, fromLast = FALSE, ...) {
  if (.use_collapse() && !fromLast && identical(incomparables, FALSE) && !is.list(x)) {
    collapse::fduplicated(x, ...)
  } else {
    duplicated(x, incomparables = incomparables, fromLast = fromLast, ...)
  }
}

f_anyDuplicated <- function(x, incomparables = FALSE, fromLast = FALSE) {
  if (.use_collapse() && !fromLast && identical(incomparables, FALSE) && !is.null(x)) {
    collapse::any_duplicated(x)
  } else {
    anyDuplicated(x, incomparables = incomparables, fromLast = fromLast)
  }
}
