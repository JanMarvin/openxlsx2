
# columns -----------------------------------------------------------------

#' Convert integer to Excel column
#'
#' Converts an integer to an Excel column label.
#'
#' @param x A numeric vector.
#' @export
#' @examples
#' int2col(1:10)
int2col <- function(x) {
  if (!is.numeric(x)) {
    stop("x must be numeric.")
  }

  sapply(x, int_to_col)
}

#' Convert Excel column to integer
#'
#' Converts an Excel column label to an integer.
#'
#' @param x A character vector
#' @return An integer column label (or `NULL` if `x` is `NULL`)
#' @export
#' @examples
#' col2int(LETTERS)
col2int <- function(x) {
  if (is.null(x)) return(NULL)

  if (is.numeric(x) || is.integer(x) || is.factor(x))
    return(as.integer(x))

  if (!is.character(x)) {
    stop("x must be character")

    if (anyNA(x))
      stop("x must be a valid character")
  }

  if (any(grepl(":", x))) {
    # loop through x until all ":" are replaced with integer sequences
    while (any(grepl(":", x))) {
      rng <- which(grepl(":", x))[1]
      spl <- strsplit(x[rng], ":")[[1]]
      # Replace "A:Z" with integer sequence. This changes the length of x
      x <- append(x[-rng], seq.int(col_to_int(spl[1]), col_to_int(spl[2])), after = rng - 1)
    }
  }

  col_to_int(x)
}


#' Return Excel cell coordinates from (x,y) coordinates
#'
#' @param cellCoords A data.frame with two columns coordinate pairs.
#' @return Excel alphanumeric cell reference
#' @examples
#' get_cell_refs(data.frame(1, 2))
#' # "B1"
#' get_cell_refs(data.frame(1:3, 2:4))
#' # "B1" "C2" "D3"
#' @noRd
get_cell_refs <- function(cellCoords) {
  assert_class(cellCoords, "data.frame")
  stopifnot(ncol(cellCoords) == 2L)

  if (!all(vapply(cellCoords, is_integer_ish, NA))) {
    stop("cellCoords must only contain integers", call. = FALSE)
  }

  l <- int2col(unlist(cellCoords[, 2]))
  paste0(l, cellCoords[, 1])
}



#' calculate the required column width
#'
#' @param base_font the base font name and fontsize
#' @param col_width column width
#' @examples
#' base_font <- wb_get_base_font(wb)
#' calc_col_width(base_font, col_width = 10)
#' @noRd
calc_col_width <- function(base_font, col_width) {

  # TODO save this instead as internal package data for quicker loading
  fw <- system.file("extdata", "fontwidth/FontWidth.csv", package = "openxlsx2")
  font_width_tab <- read.csv(fw)

  # TODO base font might not be the font used in this column
  font <- base_font$name$val
  size <- as.integer(base_font$size$val)

  sel <- font_width_tab$FontFamilyName == font & font_width_tab$FontSize == size
  # maximum digit width of selected font
  mdw <- font_width_tab$Width[sel]

  # formula from openxml.spreadsheet.column documentation. The formula returns exactly the expected
  # value, but the output in excel is still off. Therefore round to create even numbers. In my tests
  # the results were close to the initial col_width sizes. Character width is still bad, numbers are
  # way larger, therefore characters cells are to wide. Not sure if we need improve this.

  # Note: cannot reproduce the exact values with MS365 on Mac. Nevertheless, these values are closer
  # to the expected widths
  widths <- trunc((as.numeric(col_width) * mdw + 5) / mdw * 256) / 256
  widths <- round(widths, 3)

  if (any(sel <- widths > getOption("openxlsx2.maxWidth", 250))) {
    widths[sel] <- getOption("openxlsx2.maxWidth", 250)
  }

  if (any(sel <- widths <= getOption("openxlsx2.minWidth", 0))) {
    widths[sel] <- getOption("openxlsx2.minWidth", 1)
  }

  widths
}
