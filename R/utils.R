is_integer_ish <- function(x) {
  if (is.integer(x)) {
    return(TRUE)
  }

  if (!is.numeric(x)) {
    return(FALSE)
  }

  all(x[!is.na(x)] %% 1 == 0)
}

`%||%` <- function(x, y) if (is.null(x)) y else x

na_to_null <- function(x) {
  lapply(x, function(i) if (is.na(i)) NULL else i)
}

