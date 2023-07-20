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
#' @inheritParams base::tempfile pattern
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
#' simple wrapper around `stringi::stri_rand_strings()`
#'
#' @inheritParams stringi::stri_rand_strings
#' @param keep_seed logical should the default seed be kept unaltered
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
# dims helpers -----------------------------------------------------------------
#' Helper functions to work with `dims`
#'
#' Internal helpers to (de)construct a dims argument from/to a row and column
#' vector. Exported for user convenience.
#'
#' @param x a dimension object "A1" or "A1:A1"
#' @param as_integer If the output should be returned as integer, (defaults to string)
#' @param row a numeric vector of rows
#' @param col a numeric or character vector of cols
#' @returns
#'   * A `dims` string for `_to_dim` i.e  "A1:A1"
#'   * A list of rows and columns for `to_rowcol`
#' @examples
#' dims_to_rowcol("A1:J10")
#' rowcol_to_dims(1:10, 1:10)
#' @name dims_helper
NULL
#' @rdname dims_helper
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
#' @rdname dims_helper
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
#' @rdname dims_helper
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

# It is inspired heavily by `rlang::arg_match(multi = TRUE)` and `base::match.arg()`
# Does not allow partial matching.
match.arg_wrapper <- function(arg, choices, several.ok = FALSE, fn_name = NULL) {
  # Check valid argument names
  # partial matching accepted
  fn_name <- fn_name %||% "function"
  # match.arg(arg, choices = choices, several.ok = several.ok)
  # Using rlang::arg_match() would remove that.
  if (!several.ok) {
    if (length(arg) != 1) {
      stop(
        "Must provide a single argument found in ", fn_name, ": ", invalid_arg_nams, "\n", "Use one of ", valid_arg_nams,
        call. = FALSE
      )
    }
  }

  invalid_args <- !arg %in% choices
  if (any(invalid_args)) {
    invalid_arg_nams <- paste0("`", arg[invalid_args], "`", collapse = ", ")
    multi <- length(invalid_arg_nams) > 0
    plural_sentence <- ifelse(multi, " is an invalid argument for ", " are invalid arguments for ")

    valid_arg_nams <- paste0("'", choices[choices != ""], "'", collapse = ", ")
    stop(
      invalid_arg_nams, plural_sentence, fn_name, ": ", "\n", "Use any of ", valid_arg_nams,
      call. = FALSE
    )
  }
}
#' Helper to specify the `dims` argument.
#'
#' @description
#' `wb_dims()` can be used to help provide the `dims` argument, in the `wb_add_*` functions.
#' It returns a Excel range (i.e. "A1:B1") or a start like "A2".
#' It can be very useful as you can specify many parameters that interact together
#' In general, you must provide named arguments. `wb_dims()` will only accept unnamed arguments
#' if it is `rows`, `cols`, for example `wb_dims(1:4, 1:2)`, that will return "A1:B4".
#'
#' `wb_dims()` can also be used with an object (a `data.frame` or a `matrix` for example.)
#' All parameters are numeric unless stated otherwise.
#' # Using `wb_dims()` alone
#'
#' * `rows` / `cols` (if you want to specify a single one, use `start_row` / `start_col`)
#' * `start_row`
#' * `start_col`
#'
#' # Using `wb_dims()` with an object
#'
#' * `x` An object (typically a `matrix` or a `data.frame`)
#' * `start_row` the starting row of `x` (The `dims` returned will be )
#' * `start_col` the starting column: a single integer, or an Excel column identifier "A", "B" etc.
#' * `rows` (Which rows? (not fully supported yet
#' * `cols` a range of column, or one of the column names of `x` (length 1 only accepted)
#' * `row_names` A logical, should include `row_names`
#' * `col_names` A logical, defaults to `TRUE` is `x` has dimensions.
#'          Using `FALSE` can be useful to apply a style or a formula to the content of `x`.
#' @return A `dims` string
#' @param ... construct dims arguments, from rows/cols vectors or objects that
#'   can be coerced to data frame
#' @examples
#' # Provide coordinates
#' wb_dims()
#' wb_dims(1, 4)
#' wb_dims(rows = 1, cols = 4)
#' wb_dims(start_row = 4)
#' wb_dims(start_col = 2)
#' wb_dims(1:4, 6:9, start_row = 5)
#' # Provide  vectors
#' wb_dims(1:10, 1:5)
#' wb_dims(rows = 1:10, cols = 1:10)
#' # provide `start_col` / `start_row`
#' wb_dims(rows = 1:10, cols = 1:10, start_row = 2)
#' wb_dims(rows = 1:10, cols = 1:10, start_col = 2)
#' # or objects
#' wb_dims(x = mtcars)
#' # dims of all the data of mtcars.
#' wb_dims(x = mtcars, col_names = FALSE)
#' @export
wb_dims <- function(...) {
  args <- list(...)
  lengt <- length(args)
  if (lengt == 0 || (lengt == 1 && is.null(args[[1]]))) {
    return("A1")
  }

  # nams cannot be NULL now
  nams <- names(args) %||% rep("", lengt)
  valid_arg_nams <- c("x", "rows", "cols", "start_row", "start_col", "row_names", "col_names")
  any_args_named <- any(nzchar(nams))
  # unused, but can be used, if we need to check if any, but not all
  # has_some_named_args <- any(!nzchar(nams)) & any(nzchar(nams))
  # Check if valid args were provided if any argument is named.
  if (any_args_named) {
    match.arg_wrapper(arg = nams, choices = c(valid_arg_nams, ""), several.ok = TRUE, fn_name = "`wb_dims()`")
  }
  # After this point, no need to search for invalid arguments!

  n_unnamed_args <- length(which(!nzchar(nams)))
  all_args_unnamed <- n_unnamed_args == lengt
  # argument dispatch / creation here.
  # All names provided, happy :)
  # Checking if valid names were provided.

  if (n_unnamed_args > 2) {
    stop("only `rows` and `cols` can be provided without names. You must name all other arguments.")
  }
  if (lengt == 1 && all_args_unnamed) {
    stop(
      "Specifying a single unnamed argument is not handled by `wb_dims()`",
      "use `x`, `start_row`/ `start_col`. You can also use `dims = NULL`"
    )
  }
  ok_if_arg1_unnamed <-
    is.atomic(args[[1]]) | any(nams %in% c("rows", "cols"))

  if (nams[1] == "" && !ok_if_arg1_unnamed) {
    stop(
      "The first argument must either be named or be a vector.",
      "Providing a single named argument must either be `start_row`, `start_col` or `x`."
    )
  }
  if (n_unnamed_args == 1 && lengt > 1 && !"rows" %in% nams) {
    message("Assuming the first unnamed argument to be `rows`.")
    nams[which(nams == "")[1]] <- "rows"
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == lengt
  }
  if (n_unnamed_args == 1 && lengt > 1 && "rows" %in% nams) {
    message("Assuming the first unnamed argument to be `cols`.")
    nams[which(nams == "")[1]] <- "cols"
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == lengt
  }

  # if 2 unnamed arguments, will be rows, cols.
  if (n_unnamed_args == 2) {
    # message("Assuming the first 2 unnamed arguments to be `rows`, `cols` resp.")
    rows_pos <- which(nams == "")[1]
    cols_pos <- which(nams == "")[2]
    nams[c(rows_pos, cols_pos)] <- c("rows", "cols")
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == lengt
  }

  has_some_unnamed_args <- any(!nzchar(nams))
  if (has_some_unnamed_args) {
    stop("Internal error, all arguments should be named after this point.")
  }

  x_present <- "x" %in% nams
  cond_acceptable_lengt_1 <- x_present || !is.null(args$start_row) || !is.null(args$start_col)

  if (lengt == 1 && !cond_acceptable_lengt_1) {
    # Providing a single argument acceptable is only  `x`
    sentence_unnamed <- ifelse(all_args_unnamed, "unnamed ", " ")
    stop(
      "Specifying a single", sentence_unnamed, "argument to `wb_dims()` is not supported.",
      "\n",
      "use any of `x`, `start_row` `start_col`. You can also use `rows` and `cols`, You can also use `dims = NULL`"
    )
  }
  cnam_null <- is.null(args$col_names)
  rnam_null <- is.null(args$row_names)


  if (!x_present) {
    if (!cnam_null || !rnam_null) {
      stop("`row_names`, and `col_names` should only be used if `x` is present.")
    }
  }
  rows_arg <- args$rows
  #
  x <- args$x
  x_has_named_dims <- inherits(x, "data.frame") | inherits(x, "matrix")
  if (x_has_named_dims && !is.null(rows_arg)) {
    is_rows_a_colname <- rows_arg %in% colnames(x)

    if (any(is_rows_a_colname)) {
      stop("`rows` is the incorrect argument. Use `cols` instead.")
    }
  }
  if (is.character(rows_arg)) {
    warning("It's preferable to specify integers indices for `rows`", "See `col2int(rows)` to find the correct index.")
  }

  rows_arg <- col2int(rows_arg)
  cols_arg <- args$cols
  x <- args$x
  x_has_named_dims <- inherits(x, "data.frame") | inherits(x, "matrix")

  # rows_and_cols_present <- all(c("rows", "cols") %in% nams)


  # Find column location id if `cols` is named.
  if (x_has_named_dims && !is.null(cols_arg)) {
    is_cols_a_colname <- cols_arg %in% colnames(x)

    if (any(is_cols_a_colname)) {
      if (length(is_cols_a_colname) != 1) {
        stop(
          "Supplying multiple column names is not supported by the `wb_dims()` helper, use the `cols`  arguments instead.",
          "\n Use a single `cols` at a time with `wb_dims()`"
        )
      }
      col_name <- cols_arg

      cols_arg <- which(colnames(x) == cols_arg)
      message("Transforming col name = '", col_name, "' to `cols = ", cols_arg, "`")
    }
  }

  if (!is.null(rows_arg)) {
    assert_class(rows_arg, class = "integer", envir = parent.frame(n = 2), arg_nm = "rows_arg")
  }

  if (!is.null(cols_arg)) {
    cols_arg <- col2int(cols_arg)
    assert_class(cols_arg, class = "integer", envir = parent.frame(n = 2), arg_nm = "cols_arg")
  }

  srow <- args$start_row %||% 1L
  srow <- srow - 1L
  scol <- col2int(args$start_col) %||% 1L
  scol <- scol - 1L
  # after this point, no assertion, assuming all elements to be acceptable



  col_names <- args$col_names
  row_names <- args$row_names

  rows_arg
  cols_arg

  if (!all(length(scol) == 1, length(srow) == 1, scol >= 0, srow >= 0)) {
    stop("Internal error. At this point scol and srow should have length 1.")
  }


  # if `!x` return early
  if (!x_present) {
    row_span <- srow + rows_arg %||% 1L
    col_span <- scol + cols_arg %||% 1L
    if (length(row_span) == 1 && length(col_span) == 1) {
      # A1
      row_start <- row_span
      col_start <- col_span
      dims <- rowcol_to_dim(row_start, col_start)
    } else {
      # A1:B2

      dims <- rowcol_to_dims(row_span, col_span)
    }
    return(dims)
  }
  # Making sure that at this point, we only cover the case for `x`

  col_names <- col_names %||% x_has_named_dims
  row_names <- row_names %||% FALSE
  if (!col_names && row_names && x_has_named_dims) {
    warning("The combination of `row_names = TRUE` and `col_names = FALSE` is not recommended.",
      "`col_names` allows to select the region that contains the data only.",
      "`row_names` = TRUE adds row numbers if the data doesn't have rownames.",
      call. = FALSE
    )
  }
  assert_class(col_names, "logical")
  assert_class(row_names, "logical")
  x <- as.data.frame(x)

  nrow_to_span <- nrow(x)
  ncol_to_span <- ncol(x)

  if (col_names && x_has_named_dims) {
    nrow_to_span <- nrow_to_span + 1
  }
  # if without column names and with named dimensions
  # We will increment the start row by 1.
  if (!col_names && x_has_named_dims) {
    srow <- srow + 1
  }

  # Adding a column when spanning.
  if (row_names) {
    ncol_to_span <- ncol_to_span + 1
  }

  # if (row_names) {
  #   scol <- scol + 1
  # }

  if (!all(scol >= 0, srow >= 0)) {
    stop("Internal error. At this point `start_col` and `start_row` should have length 1.")
  }

  if (is.null(cols_arg) && is.null(rows_arg)) {
    # wb_dims(data.frame())
    row_span <- srow + seq_len(nrow_to_span)
    col_span <- scol + seq_len(ncol_to_span)
  } else if (is.null(rows_arg)) {
    row_span <- srow + seq_len(nrow_to_span)
    col_span <- scol + cols_arg + row_names
  } else if (is.null(cols_arg)) {
    row_span <- srow + rows_arg + col_names
    col_span <- scol + seq_len(ncol_to_span)
  } else {
    "problem"
  }

  if (length(row_span) == 1 && length(col_span) == 1) {
    # A1
    row_start <- row_span
    col_start <- col_span
    dims <- rowcol_to_dim(row_start, col_start)
  } else {
    # A1:B2

    dims <- rowcol_to_dims(row_span, col_span)
  }

  dims
}

# Relationship helpers --------------------
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
# `fmt_txt()` ------------------------------------------------------------------
#' format strings independent of the cell style.
#'
#'
#' @details
#' The result is an xml string. It is possible to paste multiple `fmt_txt()`
#' strings together to create a string with differing styles.
#'
#' Using `fmt_txt(charset = 161)` will give the Greek Character Set
#'
#'  | charset| "Character Set"      |
#'  |--------|----------------------|
#'  |  0     | "ANSI_CHARSET"       |
#'  |  1     | "DEFAULT_CHARSET"    |
#'  |  2     | "SYMBOL_CHARSET"     |
#'  | 77     | "MAC_CHARSET"        |
#'  | 128    | "SHIFTJIS_CHARSET"   |
#'  | 129    | "HANGUL_CHARSET"     |
#'  | 130    | "JOHAB_CHARSET"      |
#'  | 134    | "GB2312_CHARSET"     |
#'  | 136    | "CHINESEBIG5_CHARSET"|
#'  | 161    | "GREEK_CHARSET"      |
#'  | 162    | "TURKISH_CHARSET"    |
#'  | 163    | "VIETNAMESE_CHARSET" |
#'  | 177    | "HEBREW_CHARSET"     |
#'  | 178    | "ARABIC_CHARSET"     |
#'  | 186    | "BALTIC_CHARSET"     |
#'  | 204    | "RUSSIAN_CHARSET"    |
#'  | 222    | "THAI_CHARSET"       |
#'  | 238    | "EASTEUROPE_CHARSET" |
#'  | 255    | "OEM_CHARSET"        |
#'
# FIXME review the `fmt_txt.Rd`
# #' @param x a string or part of a string
#' @param bold bold
#' @param italic italic
#' @param underline underline
#' @param strike strike
#' @param size the font size
#' @param color a wbColor color for the font
#' @param font the font name
#' @param charset integer value from the table below
#' @param outline TRUE or FALSE
#' @param vert_align baseline, superscript, or subscript
#' @examples
#' fmt_txt("bar", underline = TRUE)
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
    vert_align = NULL
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
  if (length(vert_align)) {
    xml_vrtln <- xml_node_create("vertAlign", xml_attributes = c("val" = as_xml_attr(vert_align)))
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

#' @method + fmt_txt
#' @param x,y an openxlsx2 fmt_txt string
#' @details You can join additional objects into fmt_txt() objects using "+". Though be aware that `fmt_txt("sum:") + (2 + 2)` is different to `fmt_txt("sum:") + 2 + 2`.
#' @examples
#' fmt_txt("foo ", bold = TRUE) + fmt_txt("bar")
#' @rdname fmt_txt
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
# FIXME review the `fmt_txt.Rd`
# #' @param x an openxlsx2 fmt_txt string
#' @param ... unused
#' @examples
#' as.character(fmt_txt(2))
#' @export
as.character.fmt_txt <- function(x, ...) {
  si_to_txt(xml_node_create("si", xml_children = x))
}

#' @rdname fmt_txt
#' @method print fmt_txt
# FIXME review the `fmt_txt.Rd`
# #' @param x an openxlsx2 fmt_txt string
#' @param ... additional arguments for default print
#' @export
print.fmt_txt <- function(x, ...) {
  message("fmt_txt string: ")
  print(as.character(x), ...)
}
