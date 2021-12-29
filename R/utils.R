is_integer_ish <- function(x) {
  if (is.integer(x)) {
    return(TRUE)
  }

  if (!is.numeric(x)) {
    return(FALSE)
  }

  all(x[!is.na(x)] %% 1 == 0)
}

naToNULLList <- function(x) {
  lapply(x, function(i) if (is.na(i)) NULL else i)
}

# useful for replacing multiple x <- paste(x, new) where the length is checked
# multiple times.  This combines all elements in ... and removes anything that
# is zero length.  Much faster than multiple if/else (re)assignments
paste_c <- function(..., sep = "", collapse = " ", unlist = FALSE) {
  x <- c(...)
  if (unlist) x <- unlist(x, use.names = FALSE)
  paste(x[nzchar(x)], sep = "", collapse = " ")
}
