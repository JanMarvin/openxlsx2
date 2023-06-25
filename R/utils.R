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
  stri_join(x[nzchar(x)], sep = sep, collapse = collapse)
}

`%||%` <- function(x, y) if (is.null(x)) y else x
`%|||%` <- function(x, y) if (length(x)) x else y

na_to_null <- function(x) {
  if (is.null(x)) {
    return(NULL)
  }

  if (isTRUE(is.na(x))) {
    return(NULL)
  }

  x
}

# opposite of %in%
`%out%` <- function(x, table) {
  match(x, table, nomatch = 0L) == 0L
}

#' helper function to create temporary directory for testing purpose
#' @param name for the temp file
#' @param macros logical if the file extension is xlsm or xlsx
#' @export
temp_xlsx <- function(name = "temp_xlsx", macros = FALSE) {
  fileext <- ifelse(macros, ".xlsm", ".xlsx")
  tempfile(pattern = paste0(name, "_"), fileext = fileext)
}

#' helper function to create temporary directory for testing purpose
#' @param pattern pattern from `base::tempfile()`
#' @keywords internal
#' @noRd
temp_dir <- function(pattern = "file") {

  tmpDir <- file.path(tempfile(pattern))
  if (dir.exists(tmpDir)) {
    unlink(tmpDir, recursive = TRUE, force = TRUE)
  }

  success <- dir.create(path = tmpDir, recursive = FALSE)
  if (!success) { # nocov start
    stop(sprintf("Failed to create temporary directory '%s'", tmpDir))
  } # nocov end

  tmpDir
}

openxlsx2_options <- function() {
  options(
    # increase scipen to avoid writing in scientific
    scipen = 200,
    OutDec = ".",
    digits = 22
  )
}

unapply <- function(x, FUN, ..., .recurse = TRUE, .names = FALSE) {
  FUN <- match.fun(FUN)
  unlist(lapply(X = x, FUN = FUN, ...), recursive = .recurse, use.names = .names)
}

reg_match0 <- function(x, pat) regmatches(x, gregexpr(pat, x, perl = TRUE))
reg_match  <- function(x, pat) regmatches(x, gregexpr(pat, x, perl = TRUE))[[1]]

apply_reg_match  <- function(x, pat) unapply(x, reg_match,  pat = pat)
apply_reg_match0 <- function(x, pat) unapply(x, reg_match0, pat = pat)

wapply <- function(x, FUN, ...) {
  FUN <- match.fun(FUN)
  which(vapply(x, FUN, FUN.VALUE = NA, ...))
}

has_chr <- function(x, na = FALSE) {
  # na controls if NA is returned as TRUE or FALSE
  vapply(nzchar(x, keepNA = !na), isTRUE, NA)
}

dir_create <- function(..., warn = TRUE, recurse = FALSE) {
  # create path and directory -- returns path
  path <- file.path(...)
  dir.create(path, showWarnings = warn, recursive = recurse)
  path
}

as_binary <- function(x) {
  # To be used within a function
  if (any(x %out% list(0, 1, FALSE, TRUE))) {
    stop(deparse(x), " must be 0, 1, FALSE, or TRUE", call. = FALSE)
  }

  as.integer(x)
}

as_xml_attr <- function(x) {

  if (is.null(x)) {
    return("")
  }

  if (inherits(x, "logical")) {
    x <- as_binary(x)
  }

  if (inherits(x, "character")) {
    return(x)
  } else {
    op <- options(OutDec = ".")
    on.exit(options(op), add = TRUE)
    return(as.character(x))
  }
}

#' random string function that does not alter the seed.
#'
#' simple wrapper around `stringi::stri_rand_strings()``
#'
#' @inheritParams stringi::stri_rand_strings
#' @param keep_seed logical should the default seed be kept unaltered
#' @keywords internal
#' @noRd
random_string <- function(n = 1, length = 16, pattern = "[A-Za-z0-9]", keep_seed = TRUE) {
  # https://github.com/ycphs/openxlsx/issues/183
  # https://github.com/ycphs/openxlsx/pull/224

  if (keep_seed) {
    seed <- get0(".Random.seed", globalenv(), mode = "integer", inherits = FALSE)

    # try to get a previous openxlsx2 seed and use this as random seed
    openxlsx2_seed <- getOption("openxlsx2_seed")

    if (!is.null(openxlsx2_seed)) {
      # found one, change the global seed for stri_rand_strings
      set.seed(openxlsx2_seed)
    }
  }

  # create the random string, this alters the global seed
  res <- stringi::stri_rand_strings(n = n, length = length, pattern = pattern)

  if (keep_seed) {
    # store the altered seed and reset the global seed
    options("openxlsx2_seed" = ifelse(is.null(openxlsx2_seed), 1L, openxlsx2_seed + 1L))
    assign(".Random.seed", seed, globalenv())
  }

  return(res)
}

#' dims helpers
#' @description Internal helpers to (de)construct a dims argument from/to a row
#'  and column vector. Exported for user convenience.
#' @name dims_helper
#' @param x a dimension object "A1" or "A1:A1"
#' @param as_integer optional if the output should be returned as integer
#' @examples
#' dims_to_rowcol("A1:J10")
#' @export
dims_to_rowcol <- function(x, as_integer = FALSE) {

  dims <- x
  if (length(x) == 1 && grepl(";", x))
    dims <- unlist(strsplit(x, ";"))

  cols_out <- NULL
  rows_out <- NULL
  for (dim in dims) {
    dimensions <- unlist(strsplit(dim, ":"))
    cols <- gsub("[[:digit:]]", "", dimensions)
    rows <- gsub("[[:upper:]]", "", dimensions)

    # if "A:B"
    if (any(rows == "")) rows[rows == ""] <- "1"

    # convert cols to integer
    cols_int <- col2int(cols)
    rows_int <- as.integer(rows)

    if (length(dimensions) == 2) {
      # needs integer to create sequence
      cols <- int2col(seq.int(min(cols_int), max(cols_int)))
      rows_int <- seq.int(min(rows_int), max(rows_int))
    }

    if (as_integer) {
      cols <- cols_int
      rows <- rows_int
    } else {
      rows <- as.character(rows_int)
    }

    cols_out <- unique(c(cols_out, cols))
    rows_out <- unique(c(rows_out, rows))
  }

  list(cols_out, rows_out)
}
#' row and col to dims
#' @param row a numeric vector of rows
#' @param col a numeric or character vector of cols
#' @noRd
rowcol_to_dim <- function(row, col) {
  # no assert for col. will output character anyways
  # assert_class(row, "numeric") - complains if integer
  col_int <- col2int(col)
  min_col <- int2col(min(col_int))
  min_row <- min(row)

  # we will always return something like "A1"
  stringi::stri_join(min_col, min_row)
}

#' @rdname dims_helper
#' @param row a numeric vector of rows
#' @param col a numeric or character vector of cols
#' @examples
#' rowcol_to_dims(1:10, 1:10)
#' @export
rowcol_to_dims <- function(row, col) {

  # no assert for col. will output character anyways
  # assert_class(row, "numeric") - complains if integer

  col_int <- col2int(col)

  min_col <- int2col(min(col_int))
  max_col <- int2col(max(col_int))

  min_row <- min(row)
  max_row <- max(row)

  # we will always return something like "A1:A1", even for single cells
  stringi::stri_join(min_col, min_row, ":", max_col, max_row)

}

#' removes entries from worksheets_rels
#' @param x character string
#' @noRd
relship_no <- function(obj, x) {
  if (length(obj) == 0) return(character())
  relship <- rbindlist(xml_attr(obj, "Relationship"))
  relship$typ <- basename(relship$Type)
  relship <- relship[relship$typ != x, ]
  df_to_xml("Relationship", relship[c("Id", "Type", "Target")])
}

#' get ids from worksheets_rels
#' @param x character string
#' @noRd
get_relship_id <- function(obj, x) {
  if (length(obj) == 0) return(character())
  relship <- rbindlist(xml_attr(obj, "Relationship"))
  relship$typ <- basename(relship$Type)
  relship <- relship[relship$typ == x, ]
  unname(unlist(relship[c("Id")]))
}

#' filename_id returns an integer vector with the file name as name
#' @param x vector of filenames
#' @noRd
filename_id <- function(x) {
  vapply(X = x,
         FUN = function(file) as.integer(gsub("\\D+", "", basename(file))),
         FUN.VALUE = NA_integer_)
}

#' filename_id returns an integer vector with the file name as name
#' @param x vector of file names
#' @noRd
read_xml_files <- function(x) {
  vapply(X = x,
         FUN = read_xml,
         pointer = FALSE,
         FUN.VALUE = NA_character_,
         USE.NAMES = FALSE)
}

#' unlist modifies names
#' @param x a cf list
#' @keywords internal
#' @noRd
un_list <- function(x) {

  names <- vapply(x, length, NA_integer_)
  nams <- NULL
  for (i in seq_along(names)) {
    nam <- rep(names(names[i]), names[i])
    nams <- c(nams, nam)
  }
  x <- unlist(x, use.names = FALSE)
  names(x) <- nams
  x
}

#' format strings independent of the cell style.
#' @details
#' The result is an xml string. It is possible to paste multiple `fmt_txt()`
#' strings together to create a string with differing styles.
#' @param x a string or part of a string
#' @param bold bold
#' @param italic italic
#' @param underline underline
#' @param strike strike
#' @param size the font size
#' @param color a wbColor color for the font
#' @param font the font name
#' @param charset integer value from the table below
#' @param outline TRUE or FALSE
#' @param vertAlign baseline, superscript, or subscript
#' @details
#'  |"Value" | "Character Set"      |
#'  |--------|----------------------|
#'  |"0"     | "ANSI_CHARSET"       |
#'  |"1"     | "DEFAULT_CHARSET"    |
#'  |"2"     | "SYMBOL_CHARSET"     |
#'  |"77"    | "MAC_CHARSET"        |
#'  |"128"   | "SHIFTJIS_CHARSET"   |
#'  |"129"   | "HANGUL_CHARSET"     |
#'  |"130"   | "JOHAB_CHARSET"      |
#'  |"134"   | "GB2312_CHARSET"     |
#'  |"136"   | "CHINESEBIG5_CHARSET"|
#'  |"161"   | "GREEK_CHARSET"      |
#'  |"162"   | "TURKISH_CHARSET"    |
#'  |"163"   | "VIETNAMESE_CHARSET" |
#'  |"177"   | "HEBREW_CHARSET"     |
#'  |"178"   | "ARABIC_CHARSET"     |
#'  |"186"   | "BALTIC_CHARSET"     |
#'  |"204"   | "RUSSIAN_CHARSET"    |
#'  |"222"   | "THAI_CHARSET"       |
#'  |"238"   | "EASTEUROPE_CHARSET" |
#'  |"255"   | "OEM_CHARSET"        |
#' @examples
#' fmt_txt("bar", underline = TRUE)
#' @name fmt_txt
#' @export
fmt_txt <- function(
    x,
    bold      = FALSE,
    italic    = FALSE,
    underline = FALSE,
    strike    = FALSE,
    size      = NULL,
    color     = NULL,
    font      = NULL,
    charset   = NULL,
    outline   = NULL,
    vertAlign = NULL
) {

  xml_b     <- NULL
  xml_i     <- NULL
  xml_u     <- NULL
  xml_strk  <- NULL
  xml_sz    <- NULL
  xml_color <- NULL
  xml_font  <- NULL
  xml_chrst <- NULL
  xml_otln  <- NULL
  xml_vrtln <- NULL

  if (bold) {
    xml_b <-  xml_node_create("b")
  }
  if (italic) {
    xml_i <-  xml_node_create("i")
  }
  if (underline) {
    xml_u <-  xml_node_create("u")
  }
  if (strike) {
    xml_strk <- xml_node_create("strike")
  }
  if (length(size)) {
    xml_sz <- xml_node_create("sz", xml_attributes = c(val = as_xml_attr(size)))
  }
  if (inherits(color, "wbColour")) {
    xml_color <- xml_node_create("color", xml_attributes = color)
  }
  if (length(font)) {
    xml_font <- xml_node_create("rFont", xml_attributes = c(val = font))
  }
  if (length(charset)) {
    xml_chrst <- xml_node_create("charset", xml_attributes = c("val" = charset))
  }
  if (length(outline)) {
    xml_otln <- xml_node_create("outline", xml_attributes = c("val" = as_xml_attr(outline)))
  }
  if (length(vertAlign)) {
    xml_vrtln <- xml_node_create("vertAlign", xml_attributes = c("val" = as_xml_attr(vertAlign)))
  }

  xml_t_attr <- if (grepl("(^\\s+)|(\\s+$)", x)) c("xml:space" = "preserve") else NULL
  xml_t <- xml_node_create("t", xml_children = x, xml_attributes = xml_t_attr)

  xml_rpr <- xml_node_create(
    "rPr",
    xml_children = c(
      xml_b,
      xml_i,
      xml_u,
      xml_strk,
      xml_sz,
      xml_color,
      xml_font,
      xml_chrst,
      xml_otln,
      xml_vrtln
    )
  )

  out <- xml_node_create(
    "r",
    xml_children = c(
      xml_rpr,
      xml_t
    )
  )
  class(out) <- c("character", "fmt_txt")
  out
}

#' @rdname fmt_txt
#' @method + fmt_txt
#' @param x an openxlsx2 fmt_txt string
#' @param y an openxlsx2 fmt_txt string
#' @details You can join additional objects into fmt_txt() objects using "+". Though be aware that `fmt_txt("sum:") + (2 + 2)` is different to `fmt_txt("sum:") + 2 + 2`.
#' @examples
#' fmt_txt("foo ", bold = TRUE) + fmt_txt("bar")
#' @export
"+.fmt_txt" <- function(x, y) {

  if (!inherits(y, "character") || !inherits(y, "fmt_txt")) {
    y <- fmt_txt(y)
  }

  z <- paste0(x, y)
  class(z) <- c("character", "fmt_txt")
  z
}

#' @rdname fmt_txt
#' @method as.character fmt_txt
#' @param x an openxlsx2 fmt_txt string
#' @param ... unused
#' @examples
#' as.character(fmt_txt(2))
#' @export
as.character.fmt_txt <- function(x, ...) {
  si_to_txt(xml_node_create("si", xml_children = x))
}

#' @rdname fmt_txt
#' @method print fmt_txt
#' @param x an openxlsx2 fmt_txt string
#' @param ... additional arguments for default print
#' @export
print.fmt_txt <- function(x, ...) {
  message("fmt_txt string: ")
  print(as.character(x), ...)
}
