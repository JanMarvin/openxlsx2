
# columns -----------------------------------------------------------------

#' @name int2col
#' @title Convert integer to Excel column
#' @description Converts an integer to an Excel column label.
#' @param x A numeric vector
#' @export
#' @examples
#' int2col(1:10)
int2col <- function(x) {
  if (!is.numeric(x)) {
    stop("x must be numeric.")
  }

  sapply(x, int_to_col)
}

#' @name col2int
#' @title Convert Excel column to integer
#' @description Converts an Excel column label to an integer.
#' @param x A character vector
#' @export
#' @examples
#' col2int(LETTERS)
col2int <- function(x) {

  if (is.numeric(x) || is.integer(x) || is.factor(x))
    return(as.numeric(x))

  if (!is.character(x)) {
    stop("x must be character")

    if (any(is.na(x)))
      stop("x must be a valid character")
  }

  col_to_int(x)
}


#' @name get_cell_refs
#' @title Return excel cell coordinates from (x,y) coordinates
#' @description Return excel cell coordinates from (x,y) coordinates
#' @param cellCoords A data.frame with two columns coordinate pairs.
#' @return Excel alphanumeric cell reference
#' @examples
#' get_cell_refs(data.frame(1, 2))
#' # "B1"
#' get_cell_refs(data.frame(1:3, 2:4))
#' # "B1" "C2" "D3"
#' @export
get_cell_refs <- function(cellCoords) {
  assert_class(cellCoords, "data.frame")
  stopifnot(ncol(cellCoords) == 2L)

  if (!all(vapply(cellCoords, is_integer_ish, NA))) {
    stop("cellCoords must only contain integers", call. = FALSE)
  }

  l <- int2col(unlist(cellCoords[, 2]))
  paste0(l, cellCoords[, 1])
}
