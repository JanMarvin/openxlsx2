
# columns -----------------------------------------------------------------

#' Convert integers to spreadsheet column notation
#'
#' @description
#' `int2col()` performs the inverse operation of [col2int()], transforming numeric
#' column indices into their corresponding spreadsheet-style character labels
#' (e.g., 1 to "A", 28 to "AB"). This is essential for converting calculated
#' indices back into a format compatible with spreadsheet cell referencing.
#'
#' @details
#' The function accepts a numeric vector and maps each integer to its positional
#' representation in a base-26 derived system. This mapping follows standard
#' spreadsheet conventions where the sequence progresses from "A" through "Z",
#' followed by "AA", "AB", and so forth.
#'
#' Validation is performed to ensure the input is numeric and finite. In accordance
#' with the Office Open XML specification used by most spreadsheet software, the
#' maximum supported column index is 16,384, which corresponds to the column
#' label "XFD". Inputs exceeding this range may result in coordinates that are
#' incompatible with standard spreadsheet applications.
#'
#' @param x A numeric vector representing the column indices to be converted.
#'
#' @return A character vector of spreadsheet column labels. Returns `NULL`
#'   if the input `x` is `NULL`.
#'
#' @section Notes:
#' * Non-integer numeric values will typically be coerced or truncated; however,
#'     infinite values will trigger an error to prevent invalid coordinate
#'     generation.
#' @seealso [col2int()]
#'
#' @examples
#' # Convert a single index
#' int2col(27)
#'
#' # Convert a sequence of indices
#' int2col(1:10)
#'
#' # Handle large column indices
#' int2col(c(702, 703, 16384))
#'
#' @export
int2col <- function(x) {
  if (is.null(x)) return(NULL)

  if (!is.numeric(x) || any(is.infinite(x))) {
    stop("x must be finite and numeric.")
  }

  ox_int_to_col(x)
}

check_range <- function(x) {
  r <- suppressWarnings(range(as.numeric(x), na.rm = TRUE))
  any(r < 1 | r > 16384)
}

#' Convert spreadsheet column notation to integers
#'
#' @description
#' `col2int()` transforms spreadsheet-style column identifiers (e.g., "A", "B", "AA")
#' into their corresponding integer indices. This utility is fundamental for
#' programmatic data manipulation, where "A" is mapped to 1, "B" to 2, and "ZZ"
#' to 702.
#'
#' @details
#' The function is designed to handle various input formats encountered during
#' spreadsheet data processing. In addition to single column labels, it supports
#' range notation using the colon operator (e.g., "A:C"). When a range is
#' detected, the function internally expands the notation into a complete
#' sequence of integers (e.g., 1, 2, 3). This behavior is particularly useful
#' when passing column selections to functions like [wb_to_df()] or [wb_read()].
#'
#' Input validation ensures that only atomic vectors are processed. If the input
#' is already numeric or a factor, the function ensures the values fall within
#' the valid spreadsheet column range before coercion to integers. Note that
#' the presence of `NA` values in the input will trigger an error to maintain
#' data integrity during index calculation.
#'
#' @param x A character vector of column labels, a numeric vector of indices,
#'   or a factor. Supports range notation like "A:Z".
#'
#' @return An integer vector representing the column indices. Returns `NULL`
#'   if the input `x` is `NULL`, or an empty integer vector if the length of
#'   `x` is zero.
#'
#' @section Notes:
#' * Range expansion via `:` is performed iteratively until all sequences are
#'     resolved into individual integer components.
#' * In compliance with spreadsheet software standards, the function validates
#'     that indices do not exceed the maximum allowable column limit.
#'
#' @seealso [int2col()]
#'
#' @examples
#' # Convert standard labels
#' col2int(c("A", "B", "Z"))
#'
#' # Convert ranges to integer sequences
#' col2int("A:C")
#'
#' # Mix individual columns and ranges
#' col2int(c("A", "C:E", "G"))
#'
#' # Handle numeric inputs
#' col2int(c(1, 2, 26))
#'
#' @export
col2int <- function(x) {
  if (is.null(x)) return(NULL)
  if (!is.atomic(x)) {
    stop("x must be character")
  }
  if (length(x) == 0) return(integer())

  if (is.numeric(x) || is.factor(x)) {
    if (check_range(x))
      stop("Column exceeds valid range", call. = FALSE)
    return(as.integer(x))
  }

  if (anyNA(x)) stop("x contains NA")

  if (any(grepl(":", x))) {
    # loop through x until all ":" are replaced with integer sequences
    while (any(grepl(":", x))) {
      rng <- grep(":", x)[1]
      spl <- strsplit(x[rng], ":")[[1]]
      # Replace "A:Z" with integer sequence. This changes the length of x
      x <- append(x[-rng], seq.int(col_to_int(spl[1]), col_to_int(spl[2])), after = rng - 1)
    }
  }

  col_to_int(x)
}

#' Converter spreadsheet row to integer
#'
#' Converts character row to integer and checks that the range is valid
#' @param x a dimension
#' @noRd
#' @examples
#' row2int("A1")
row2int <- function(x) {

  if (is.null(x)) return(NULL)
  if (length(x) == 0) return(integer())

  rows <- as.integer(chartr("ABCDEFGHIJKLMNOPQRSTUVWXYZ",
                            "                          ", x))

  if (anyNA(rows)) stop("missings not allowed in rows", call. = FALSE)

  rr <- range(rows, na.rm = TRUE)

  if (any(rr < 1L | rr > 1048576L))
    stop("Row exceeds valid range", call. = FALSE)

  rows
}


## TODO replace with `wb_dims(rows = ..., cols = ...)`
#' Return spreadsheet cell coordinates from (x,y) coordinates
#'
#' @param cellCoords A data.frame with two columns coordinate pairs.
#' @return spreadsheet alphanumeric cell reference
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

  if (!any(font %in% font_width_tab$FontFamilyName))
    font <- "Aptos Narrow"

  sel <- font_width_tab$FontFamilyName == font & font_width_tab$FontSize == size
  # maximum digit width of selected font
  mdw <- font_width_tab$Width[sel]

  # formula from openxml.spreadsheet.column documentation. The formula returns exactly the expected
  # value, but the output in spreadsheet software is still off. Therefore round to create even numbers.
  # In my tests the results were close to the initial col_width sizes. Character width is still bad,
  # numbers are way larger, therefore characters cells are to wide. Not sure if we need improve this.

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
