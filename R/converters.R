
# columns -----------------------------------------------------------------

#' @name int2col
#' @title Convert integer to Excel column
#' @description Converts an integer to an Excel column label.
#' @param x A numeric vector
#' @export
#' @examples
#' int2col(1:10)
int2col <- function(x) {
  od <- getOption("OutDec")
  options("OutDec" = ".")
  on.exit(expr = options("OutDec" = od), add = TRUE)

  if (!is.numeric(x)) {
    stop("x must be numeric.")
  }

  convert_to_excel_ref(cols = x, LETTERS = LETTERS)
}

#' @name col2int
#' @title Convert Excel column to integer
#' @description Converts an Excel column label to an integer.
#' @param x A character vector
#' @export
#' @examples
#' col2int(LETTERS)
col2int <- function(x) {

  if (!is.character(x)) {
    stop("x must be character")

    if (any(is.na(x)))
      stop("x must be a valid character")
  }

  convert_from_excel_ref(x)
}


# refs --------------------------------------------------------------------

#' @name convertFromExcelRef
#' @title Convert excel column name to integer index
#' @description Convert excel column name to integer index e.g. "J" to 10
#' @param col An excel column reference
#' @export
#' @examples
#' convertFromExcelRef("DOG")
#' convertFromExcelRef("COW")
#'
#' ## numbers will be removed
#' convertFromExcelRef("R22")
convertFromExcelRef <- function(col) {

  ## increase scipen to avoid writing in scientific
  exSciPen <- getOption("scipen")
  od <- getOption("OutDec")
  options("scipen" = 10000)
  options("OutDec" = ".")

  on.exit(options("scipen" = exSciPen), add = TRUE)
  on.exit(expr = options("OutDec" = od), add = TRUE)

  col <- toupper(col)
  charFlag <- grepl("[A-Z]", col)
  if (any(charFlag)) {
    col[charFlag] <- gsub("[0-9]", "", col[charFlag])
    d <- lapply(strsplit(col[charFlag], split = ""), function(x) match(rev(x), LETTERS))
    col[charFlag] <- unlist(lapply(seq_along(d), function(i) sum(d[[i]] * (26^(0:(length(d[[i]]) - 1L))))))
  }

  col[!charFlag] <- as.integer(col[!charFlag])

  return(as.integer(col))
}


#' @name getCellRefs
#' @title Return excel cell coordinates from (x,y) coordinates
#' @description Return excel cell coordinates from (x,y) coordinates
#' @param cellCoords A data.frame with two columns coordinate pairs.
#' @return Excel alphanumeric cell reference
#' @examples
#' getCellRefs(data.frame(1, 2))
#' # "B1"
#' getCellRefs(data.frame(1:3, 2:4))
#' # "B1" "C2" "D3"
#' @export
getCellRefs <- function(cellCoords) {
  assert_class(cellCoords, "data.frame")
  stopifnot(ncol(cellCoords) == 2L)

  if (!all(vapply(cellCoords, is_integer_ish, NA))) {
    stop("cellCoords must only contain integers", call. = FALSE)
  }

  od <- getOption("OutDec")
  options("OutDec" = ".")
  on.exit(expr = options("OutDec" = od), add = TRUE)

  l <- convert_to_excel_ref(cols = unlist(cellCoords[, 2]), LETTERS = LETTERS)
  paste0(l, cellCoords[, 1])
}


# others ------------------------------------------------------------------

convert2EMU <- function(d, units) {
  if (grepl("in", units)) {
    d <- d * 2.54
  }

  if (grepl("mm|milli", units)) {
    d <- d / 10
  }

  return(d * 360000)
}


pixels2ExcelColWidth <- function(pixels) {
  if (any(!is.numeric(pixels))) {
    stop("All elements of pixels must be numeric")
  }

  pixels[pixels == 0] <- 8.43
  pixels[pixels != 0] <- (pixels[pixels != 0] - 12) / 7 + 1

  pixels
}

