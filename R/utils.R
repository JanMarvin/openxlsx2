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
  if (length(x) == 1 && is.na(x)) x <- c(NA, NA, NA)
  lapply(x, function(i) if (is.na(i)) NULL else i)
}

# useful for replacing multiple x <- paste(x, new) where the length is checked
# multiple times.  This combines all elements in ... and removes anything that
# is zero length.  Much faster than multiple if/else (re)assignments
paste_c <- function(..., sep = "", collapse = " ") {
  x <- c(...)
  stringi::stri_join(x[nzchar(x)], sep = sep, collapse = collapse)
}

if (getRversion() < "4.4.0") {
  `%||%` <- function(x, y) if (is.null(x)) y else x
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

  base <- tempdir()
  repeat {
    path <- file.path(base, paste0(pattern, "_", basename(tempfile())))
    if (dir.create(path, showWarnings = FALSE)) return(path)
  }
}

default_save_opt <- function() {
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

wapply <- function(x, FUN, ...) {
  FUN <- match.fun(FUN)
  which(vapply(x, FUN, FUN.VALUE = NA, ...))
}

dir_create <- function(..., warn = TRUE, recurse = FALSE) {
  # create path and directory -- returns path
  path <- file.path(...)
  dir.create(path, showWarnings = warn, recursive = recurse)
  path
}

as_binary <- function(x) {
  # To be used within a function
  if (any(!x %in% list(0, 1, FALSE, TRUE))) {
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
    x
  } else {
    op <- options(OutDec = ".")
    on.exit(options(op), add = TRUE)
    as.character(x)
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

  res
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
#' @param single argument indicating if [rowcol_to_dims()] returns a single cell dimension
#' @param fix setting the type of the reference. Per default, no type is set. Options are
#' `"all"`, `"row"`, and `"col"`
#' @returns
#'   * A `dims` string for `_to_dim` i.e  "A1:A1"
#'   * A named list of rows and columns for `to_rowcol`
#' @examples
#' dims_to_rowcol("A1:J10")
#' wb_dims(1:10, 1:10)
#'
#' @name dims_helper
#' @seealso [wb_dims()]
NULL
#' @rdname dims_helper
#' @export
dims_to_rowcol <- function(x, as_integer = FALSE) {

  dims <- x
  if (length(x) == 1 && inherits(x, "character")) {
    if (any(grepl(",|;", x))) dims <- unlist(strsplit(dims, split = "[,;]"))
  }

  cols_out <- NULL
  rows_out <- NULL
  for (dim in dims) {
    dimensions <- unlist(strsplit(dim, ":"))
    dimensions <- gsub("\\$", "", dimensions)
    cols <- gsub("[[:digit:]]", "", dimensions)
    rows <- gsub("[[:upper:]]", "", dimensions)

    # if "A:B"
    # FIXME this seems poorly handled. A:B should return c(1, 1041048576)
    if (any(rows == "")) rows[rows == ""] <- "1"

    # FIXME "1:1" seems to be entirely unhandled
    if (any(cols == "")) {
      stop(
        "A dims string passed to `dims_to_rowcol()` contained no alphabetic column",
        call. = FALSE
      )
    }

    # convert cols to integer
    cols_int <- col2int(cols)
    rows_int <- row2int(rows)

    if (length(dimensions) == 2) {
      # needs integer to create sequence
      cols <- int2col(seq.int(cols_int[1], cols_int[2]))
      rows_int <- seq.int(rows_int[1], rows_int[2])
    }

    if (as_integer) {
      cols <- cols_int
      rows <- rows_int
    } else {
      rows <- as.character(rows_int)
    }

    cols_out <- union(cols_out, cols)
    rows_out <- union(rows_out, rows)
  }

  list(col = cols_out, row = rows_out)
}


#' @rdname dims_helper
#' @export
validate_dims <- function(x) {
  assert_class(x, "character")
  dims <- gsub("\\$", "", x)

  if (any(dims == ""))
    stop("Unexpected blank strings in dims validation detected", call. = FALSE)

  dims <- unlist(strsplit(dims, "[,;]"))

  if (any(grepl("[^A-Z0-9:]", dims)))
    stop("dims contains invalid character", call. = FALSE)

  dm_list <- strsplit(dims, ":")

  all_parts <- unlist(dm_list)

  has_alpha <- grepl("[[:alpha:]]", all_parts)
  has_digit <- grepl("[[:digit:]]", all_parts)

  cols <- if (any(has_alpha)) col2int(all_parts) else c(1, 16384)
  rows <- if (any(has_digit)) row2int(all_parts) else c(1, 1048576)

  # should be TRUE, otherwise the functions above would have thrown an error
  all(is.numeric(cols) & is.numeric(rows))
}


#' @rdname dims_helper
#' @noRd
rowcol_to_dim <- function(row, col, fix = NULL) {
  # no assert for col. will output character anyways
  # assert_class(row, "numeric") - complains if integer
  col_int <- col2int(col)
  min_col <- int2col(min(col_int))
  min_row <- min(row)

  # we will always return something like "A1"
  if (!is.null(fix)) {
    match.arg(fix, c("all", "col", "row", "none"))
    if (fix == "all")
      return(stringi::stri_join("$", min_col, "$", min_row))
    if (fix == "col")
      return(stringi::stri_join("$", min_col, min_row))
    if (fix == "row")
      return(stringi::stri_join(min_col, "$", min_row))
  }

  stringi::stri_join(min_col, min_row)
}


# begin - end
rc_to_dims <- function(cb, rb, ce, re, fix = NULL) {
  if (is.null(fix) || length(fix) == 1) {
    sell <- 1
    selr <- 1
  } else if (length(fix) == 2) {
    sell <- 1
    selr <- 2
  }
  fixl <- fix[sell]
  fixr <- fix[selr]

  stringi::stri_join(rowcol_to_dim(rb, cb, fixl), ":", rowcol_to_dim(re, ce, fixr))
}

#' @rdname dims_helper
#' @export
rowcol_to_dims <- function(row, col, single = TRUE, fix = NULL) {

  # no assert for col. will output character anyways
  # assert_class(row, "numeric") - complains if integer

  row     <- as.integer(row)
  col_int <- col2int(col)

  if (col_int[1] < col[length(col_int)]) {
    min_col <- min(col_int)
    max_col <- max(col_int)
  } else {
    min_col <- max(col_int)
    max_col <- min(col_int)
  }

  if (row[1] < row[length(row)]) {
    min_row <- min(row)
    max_row <- max(row)
  } else {
    min_row <- max(row)
    max_row <- min(row)
  }

  # we will always return something like "A1:A1", even for single cells
  if (single) {
    rc_to_dims(min_col, min_row, max_col, max_row, fix = fix)
  } else {
    paste0(vapply(int2col(col_int), FUN = function(x) rc_to_dims(x, min_row, x, max_row, fix = fix), ""), collapse = ",")
  }

}

#' consecutive range in vector
#' @param x integer vector
#' @keywords internal
con_rng <- function(x) {
  if (length(x) == 0) return(NULL)
  group <- cumsum(c(1, diff(x) != 1))

  # Extract the first and last element of each group using tapply
  ranges <- tapply(x, group, function(y) c(beg = y[1], end = y[length(y)]))
  ranges_df <- do.call(rbind, ranges)

  as.data.frame(ranges_df, stringsAsFactors = FALSE)
}

#' create consecutive dims from column and row vector
#' @param cols,rows integer vectors
#' @keywords internal
con_dims <- function(cols, rows) {

  c_cols <- con_rng(cols)
  c_rows <- con_rng(rows)

  c_cols$beg <- int2col(c_cols$beg)
  c_cols$end <- int2col(c_cols$end)

  dims_cols <- paste0(c_cols$beg, "%s:", c_cols$end, "%s")

  out <- NULL
  for (i in seq_along(dims_cols)) {
    for (j in seq_len(nrow(c_rows))) {
      beg_row <- c_rows[j, "beg"]
      end_row <- c_rows[j, "end"]

      dims <- sprintf(dims_cols[i], beg_row, end_row)
      out <- c(out, dims)
    }
  }

  paste0(out, collapse = ",")
}

check_wb_dims_args <- function(args, select = NULL) {
  select <- match.arg(select, c("x", "data", "col_names", "row_names"))

  cond_acceptable_len_1 <- !is.null(args$from_row) || !is.null(args$from_col) || !is.null(args$x) || !is.null(args$from_dims)
  nams <- names(args) %||% rep("", length(args))
  all_args_unnamed <- !any(nzchar(nams))

  if (length(args) == 1 && !cond_acceptable_len_1) {
    # Providing a single argument acceptable is only  `x`
    sentence_unnamed <- ifelse(all_args_unnamed, " unnamed ", " ")
    stop(
      "Supplying a single", sentence_unnamed, "argument to `wb_dims()` is not supported. \n",
      "Use any of `x`, `from_dims`, `from_row` `from_col`. You can also use `rows` and `cols`, or `dims = NULL`",
      call. = FALSE
    )
  }
  cnam_null <- is.null(args$col_names)
  rnam_null <- is.null(args$row_names)
  if (is.character(args$rows) || is.character(args$from_row)) {
    warning("`rows` and `from_row` in `wb_dims()` should not be a character. Please supply an integer vector.", call. = FALSE)
  }

  if (is.null(args$x)) {
    if (!cnam_null || !rnam_null) {
      stop("In `wb_dims()`, `row_names`, and `col_names` should only be used if `x` is present.", call. = FALSE)
    }
  }

  x_has_colnames <- !is.null(colnames(args$x))

  if (is.character(args$cols) && x_has_colnames) {
    # Checking whether cols is character, and error if it is not the col names of x
    missing_cols <- args$cols[!(args$cols %in% colnames(args$x))]

    if (length(missing_cols) > 0) {
      missing_label <- paste0("`", missing_cols, "`", collapse = ", ")

      stop(
        "`cols` must be an integer or an existing column name of `x`. \n",
        "The following were not found: ", missing_label,
        call. = FALSE
      )
    }
  }

  invisible(NULL)
}

# it is a wrapper around base::match.arg(), but it doesn't allow partial matching.
# It also provides a more informative error message in case it fails.
match.arg_wrapper <- function(arg,
                              choices,
                              fn_name = NULL,
                              arg_name = NULL) {
  # Check valid argument names
  # partial matching accepted
  fn_name <- fn_name %||% "fn_name"

  invalid_args <- arg[arg != "" & !arg %in% choices]

  if (length(invalid_args) > 0) {
    valid_arg_nams <- paste0("'", choices[choices != ""], "'", collapse = ", ")

    invalid_labels <- paste0("`", invalid_args, "`", collapse = ", ")

    if (is.null(arg_name)) {
      msg_start <- ifelse(length(invalid_args) > 1, " are invalid arguments", " is an invalid argument")
      stop(
        invalid_labels, msg_start, " for `", fn_name, "()`\n",
        "Use any of ", valid_arg_nams,
        call. = FALSE
      )
    } else {
      msg_start <- ifelse(length(invalid_args) > 1, " are invalid values", " is an invalid value")
      stop(
        invalid_labels, msg_start, " for `", arg_name, "` in `", fn_name, "()`\n",
        "Use any of ", valid_arg_nams,
        call. = FALSE
      )
    }
  }
  arg
}

# Returns the correct select value, based on input.
# By default, it will be "data' when `x` is provided
# It will be the value if `rows` or `cols` is provided.
# It will be whatever was provided, if `select` is provided.
# But this function checks if the input is valid.
# only check WHICH arguments are provided, not what was provided.
determine_select_valid <- function(args, select = NULL) {
  valid_cases <- list(
    "x" = TRUE,
    "col_names" = !is.null(args$x) && (isTRUE(args$col_names) || is.null(args$col_names)) && is.null(args$rows),
    "row_names" = !is.null(args$x) && isTRUE(args$row_names) && is.null(args$cols), # because default is FALSE
    "data" = TRUE
  )

  default_select <- if (isFALSE(args$col_names) || !is.null(args$rows) || !is.null(args$cols)) {
    "data"
  } else {
    "x"
  }

  select <- select %||% default_select

  # Validate that the string is one of the 4 allowed names
  match.arg_wrapper(
    select,
    choices = names(valid_cases),
    fn_name = "wb_dims",
    arg_name = "select"
  )

  # Check if the specific case is valid for the current input x
  if (isFALSE(valid_cases[[select]])) {
      # If the default for row_names ever changes in openxlsx2, this would need adjustment.
    if (identical(select, "row_names")) {
      stop(
        "`select` can't be \"row_names\" if `x` doesn't have row names.\n",
        "Use `row_names = TRUE` inside `wb_dims()` to ensure row names are preserved.",
        call. = FALSE
      )
      # If the default for col_names ever changes in openxlsx2, this would need adjustment.
    } else if (identical(select, "col_names")) {
      stop(
        "`select` can't be \"col_names\" if `x` doesn't have column names.\n",
        "Use `col_names = TRUE` inside `wb_dims()` to ensure column names are preserved.",
        call. = FALSE
      )
    }
    # Removed the 'else' stop as it is unreachable
  }

  select
}

#' Helper to specify the `dims` argument
#'
#' @description
#'
#' `wb_dims()` can be used to help provide the `dims` argument, in the `wb_add_*` functions.
#' It returns a A1 spreadsheet range ("A1:B1" or "A2").
#' It can be very useful as you can specify many parameters that interact together
#' In general, you must provide named arguments. `wb_dims()` will only accept unnamed arguments
#' if they are `rows`, `cols`, for example `wb_dims(1:4, 1:2)`, that will return "A1:B4".
#'
#' `wb_dims()` can also be used with an object (a `data.frame` or a `matrix` for example.)
#' All parameters are numeric unless stated otherwise.
#'
#' # Using `wb_dims()` without an `x` object
#'
#' * `rows` / `cols` (if you want to specify a single one, use `from_row` / `from_col`)
#' * `from_row` / `from_col` the starting position of the `dims`
#'    (similar to `start_row` / `start_col`, but with a clearer name.)
#'
#' # Using `wb_dims()` with an `x` object
#'
#' `wb_dims()` with an object has 8 use-cases (they work with any position values of `from_row` / `from_col`),
#' `from_col/from_row` correspond to the coordinates at the top left of `x` including column and row names if present.
#'
#' These use cases are provided without `from_row / from_col`, but they work also with `from_row / from_col`.
#'
#' 1. provide the full grid with `wb_dims(x = mtcars)`
#' 2. provide the data grid `wb_dims(x = mtcars, select = "data")`
#' 3. provide the `dims` of column names `wb_dims(x = mtcars, select = "col_names")`
#' 4. provide the `dims` of row names  `wb_dims(x = mtcars, row_names = TRUE, select = "row_names")`
#' 5. provide the `dims` of a row span `wb_dims(x = mtcars, rows = 1:10)` selects
#'    the first 10 data rows of `mtcars` (ignoring column names)
#' 6. provide the `dims` of the data in a column span `wb_dims(x = mtcars, cols = 1:5)`
#'    select the data first 5 columns of `mtcars`
#' 7. provide a column span (including column names) `wb_dims(x = mtcars, cols = 4:7, select = "x")`
#'    select the data columns 4, 5, 6, 7 of `mtcars` + column names
#' 8. provide the position of a single column by name `wb_dims(x = mtcars, cols = "mpg")`.
#' 9. provide a row span with a column. `wb_dims(x = mtcars, cols = "mpg", rows = 5:22)`
#'
#' To reuse, a good trick is to create a wrapper function, so that styling can be
#' performed seamlessly.
#'
#' ```R
#' wb_dims_cars <- function(...) {
#'   wb_dims(x = mtcars, from_row = 2, from_col = "B", ...)
#' }
#' # using this function
#' wb_dims_cars()                     # full grid (data + column names)
#' wb_dims_cars(select = "data")      # data only
#' wb_dims_cars(select = "col_names") # select column names
#' wb_dims_cars(cols = "vs")          # select the `vs` column
#' ```
#'
#' It can be very useful to apply many rounds of styling sequentially.
#'
#'
#' @details
#'
#' When using `wb_dims()` with an object, the default behavior is
#' to select only the data / row or columns in `x`
#' If you need another behavior, use `wb_dims()` without supplying `x`.
#'
#' * `x` An object (typically a `matrix` or a `data.frame`, but a vector is also accepted.)
#' * `from_row` / `from_col` / `from_dims` the starting position of `x`
#'   (The `dims` returned will assume that the top left corner of `x` is at `from_row / from_col`
#' * `rows` Optional Which row span in `x` should this apply to.
#'   If `rows` = 0, only column names will be affected.
#' * `cols` a range of columns id in `x`, or one of the column names of `x`
#'   (length 1 only accepted for column names of `x`.)
#' * `row_names` A logical, this is to let `wb_dims()` know that `x` has row names or not.
#'   If `row_names = TRUE`, `wb_dims()` will increment `from_col` by 1.
#' * `col_names` `wb_dims()` assumes that if `x` has column names, then trying to find the `dims`.
#'
#'
#' `wb_dims()` tries to support most possible cases with `row_names = TRUE` and `col_names = FALSE`,
#' but it works best if `x` has named dimensions (`data.frame`, `matrix`), and those parameters are not specified.
#'  data with column names, and without row names. as the code is more clean.
#'
#' In the `add_data()` / `add_font()` example, if writing the data with row names
#'
#' While it is possible to construct dimensions from decreasing rows and columns, the output will always order the rows top to bottom. So
#' `wb_dims(rows = 3:1, cols = 3:1)` will not result in `"C3:A1"` and if passed to functions, it will return the same as `"C1:A3"`.
#'
#' @param ... construct `dims` arguments, from rows/cols vectors or objects that
#'   can be coerced to data frame. `x`, `rows`, `cols`, `from_row`, `from_col`, `from_dims`
#'   `row_names`, and `col_names` are accepted.
#' @param select A string, one of the followings.
#'    it improves the selection of various parts of `x`
#'    One of "x", "data", "col_names", or "row_names".
#'    `"data"` will only select the data part, excluding row names and column names  (default if `cols` or `rows` are specified)
#'    `"x"` Includes the data, column and row names if they are present. (default if none of `rows` and `cols` are provided)
#'    `"col_names"` will only return column names
#'    `"row_names"` Will only return row names.
#'
#' @return A `dims` string
#' @export
#' @examples
#' # Provide coordinates
#' wb_dims(1, 4)
#' wb_dims(rows = 1, cols = 4)
#' wb_dims(from_row = 4)
#' wb_dims(from_col = 2)
#' wb_dims(from_col = "B")
#' wb_dims(1:4, 6:9, from_row = 5)
#' # Provide vectors
#' wb_dims(1:10, c("A", "B", "C"))
#' wb_dims(rows = 1:10, cols = 1:10)
#' # provide `from_col` / `from_row`
#' wb_dims(rows = 1:10, cols = c("A", "B", "C"), from_row = 2)
#' wb_dims(rows = 1:10, cols = 1:10, from_col = 2)
#' wb_dims(rows = 1:10, cols = 1:10, from_dims = "B1")
#' # or objects
#' wb_dims(x = mtcars, col_names = TRUE)
#'
#' # select all data
#' wb_dims(x = mtcars, select = "data")
#'
#' # column names of an object (with the special select = "col_names")
#' wb_dims(x = mtcars, select = "col_names")
#'
#'
#' # dims of the column names of an object
#' wb_dims(x = mtcars, select = "col_names", col_names = TRUE)
#'
#' ## add formatting to `mtcars` using `wb_dims()`----
#' wb <- wb_workbook()
#' wb$add_worksheet("test wb_dims() with an object")
#' dims_mtcars_and_col_names <- wb_dims(x = mtcars)
#' wb$add_data(x = mtcars, dims = dims_mtcars_and_col_names)
#'
#' # Put the font as Arial for the data
#' dims_mtcars_data <- wb_dims(x = mtcars, select = "data")
#' wb$add_font(dims = dims_mtcars_data, name = "Arial")
#'
#' # Style col names as bold using the special `select = "col_names"` with `x` provided.
#' dims_column_names <- wb_dims(x = mtcars, select = "col_names")
#' wb$add_font(dims = dims_column_names, bold = TRUE, size = 13)
#'
#' # Finally, to add styling to column "cyl" (the 4th column) (only the data)
#' # there are many options, but here is the preferred one
#' # if you know the column index, wb_dims(x = mtcars, cols = 4) also works.
#' dims_cyl <- wb_dims(x = mtcars, cols = "cyl")
#' wb$add_fill(dims = dims_cyl, color = wb_color("pink"))
#'
#' # Mark a full column as important(with the column name too)
#' wb_dims_vs <- wb_dims(x = mtcars, cols = "vs", select = "x")
#' wb$add_fill(dims = wb_dims_vs, fill = wb_color("yellow"))
#' wb$add_conditional_formatting(dims = wb_dims(x = mtcars, cols = "mpg"), type = "dataBar")
#' # wb_open(wb)
#'
#' # fix relative ranges
#' wb_dims(x = mtcars) # equal to none: A1:K33
#' wb_dims(x = mtcars, fix = "all") # $A$1:$K$33
#' wb_dims(x = mtcars, fix = "row") # A$1:K$33
#' wb_dims(x = mtcars, fix = "col") # $A1:$K33
#' wb_dims(x = mtcars, fix = c("col", "row")) # $A1:K$33
wb_dims <- function(..., select = NULL) {
  args <- list(...)
  len <- length(args)

  if (len == 0 || (len == 1 && is.null(args[[1]]))) {
    stop("`wb_dims()` requires any of `rows`, `cols`, `from_row`, `from_col`, `from_dims`, or `x`.", call. = FALSE)
  }

  # nams cannot be NULL now
  nams <- names(args) %||% rep("", len)
  valid_arg_nams <- c("x", "rows", "cols", "from_row", "from_col", "from_dims", "row_names", "col_names",
                      "left", "right", "above", "below", "select", "fix")
  any_args_named <- any(nzchar(nams))
  # unused, but can be used, if we need to check if any, but not all
  # Check if valid args were provided if any argument is named.
  if (any_args_named) {
    if (any(c("start_col", "start_row") %in% nams)) {
      stop("Use `from_row` / `from_col` instead of `start_row` / `start_col`")
    }
    match.arg_wrapper(arg = nams, choices = c(valid_arg_nams, ""), fn_name = "wb_dims")
  }
  # After this point, no need to search for invalid arguments!

  n_unnamed_args <- length(which(!nzchar(nams)))
  all_args_unnamed <- n_unnamed_args == len
  # argument dispatch / creation here.
  # All names provided, happy :)
  # Checking if valid names were provided.

  if (n_unnamed_args > 2) {
    stop(
      "Only `rows` and `cols` can be provided unnamed to `wb_dims()`.\n",
      "You must name all other arguments.",
      call. = FALSE
    )
  }

  if (len == 1 && all_args_unnamed) {
    stop(
      "Supplying a single unnamed argument is not handled by `wb_dims()`.\n",
      "Use `x`, `from_row` / `from_col`.",
      call. = FALSE
    )
  }

  ok_if_arg1_unnamed <-
    is.null(args[[1]]) || is.atomic(args[[1]]) || any(nams %in% c("rows", "cols"))

  if (nams[1] == "" && !ok_if_arg1_unnamed) {
    stop(
      "The first argument must either be named or be a vector.",
      "Providing a single named argument must either be `from_dims` `from_row`, `from_col` or `x`."
    )
  }

  rows_arg <- args$rows
  cols_arg <- args$cols

  if (anyNA(rows_arg) || anyNA(cols_arg)) {
    stop("NAs are not supported in wb_dims()")
  }

  if ((!is.null(rows_arg) && !is.atomic(rows_arg)) ||
      (!is.null(cols_arg) && !is.atomic(cols_arg))) {
    stop("Input must be a vector type")
  }

  if (is.factor(rows_arg) || is.factor(cols_arg)) {
    stop("factors are not supported in wb_dims()")
  }

  if (n_unnamed_args == 1 && len > 1 && !"rows" %in% nams) {
    message("Assuming the first unnamed argument to be `rows`.")
    rows_pos <- which(nams == "")[1]
    nams[rows_pos] <- "rows"
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == len

    rows_arg <- args[[rows_pos]]
  }

  if (n_unnamed_args == 1 && len > 1 && "rows" %in% nams) {
    message("Assuming the first unnamed argument to be `cols`.")
    cols_pos <- which(nams == "")[1]
    nams[cols_pos] <- "cols"
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == len

    cols_arg <- args[[cols_pos]]
  }

  # if 2 unnamed arguments, will be rows, cols.
  if (n_unnamed_args == 2) {
    # message("Assuming the first 2 unnamed arguments to be `rows`, `cols` resp.")
    rows_pos <- which(nams == "")[1]
    cols_pos <- which(nams == "")[2]
    nams[c(rows_pos, cols_pos)] <- c("rows", "cols")
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == len

    rows_arg <- args[[rows_pos]]
    cols_arg <- args[[cols_pos]]
  }

  if (!is.null(cols_arg)) {
    is_lwr_one <- FALSE
    # cols_arg could be name(s) in x or must indicate a positive integer
    if (is.null(args$x) || (!is.null(args$x) && !all(cols_arg %in% names(args$x))))
      is_lwr_one <- if (is.numeric(cols_arg))
          min(cols_arg) < 1L
        else
          min(col2int(cols_arg)) < 1L

    if (is_lwr_one)
      stop("You must supply positive values to `cols`")
  }

  # in case the user mixes up column and row
  if (is.character(rows_arg) && !is.null(args$x)) {
    is_rows_a_colname <- rows_arg %in% colnames(args$x)
    if (any(is_rows_a_colname)) {
      stop(
        "`rows` is the incorrect argument in this case\n",
        "Use `cols` instead. Subsetting rows by name is not supported.",
        call. = FALSE
      )
    }
  }

  if (is.character(rows_arg) && !all(is_charnum(rows_arg))) {
    stop("`rows` is character and contains nothing that can be interpreted as number.", call. = FALSE)
  }

  if (!is.null(rows_arg) && (min(as.integer(rows_arg)) < 1L)) {
    stop("You must supply positive values to `rows`.")
  }

  if (length(args$from_col) > 1 || length(args$from_row) > 1 ||
    (!is.null(args$from_row) && is.character(args$from_row) && !is_charnum(args$from_row))) {
    stop("from_col/from_row must be positive integers if supplied.")
  }

  # handle from_dims
  if (!is.null(args$from_dims)) {
    if (!is.null(args$from_col) || !is.null(args$from_row)) {
      stop("Can't handle `from_row` and `from_col` if `from_dims` is supplied.")
    }
    # transform to
    from_row_and_col <- dims_to_rowcol(args$from_dims, as_integer = TRUE)
    fcol <- from_row_and_col[["col"]]
    frow <- from_row_and_col[["row"]]
  } else {
    fcol <- args$from_col %||% 1L
    frow <- args$from_row %||% 1L
  }

  fcol <- col2int(fcol)
  frow <- col2int(frow)

  left  <- args$left
  right <- args$right
  above <- args$above
  below <- args$below


  # there can be only one
  if (length(c(left, right, above, below)) > 1) {
    stop("can only be one direction")
  }

  # default is column names and no row names
  cnms <- args$col_names %||% 1
  rnms <- args$row_names %||% 0

  # NCOL(NULL)/NROW(NULL) could work as well, but the behavior might have
  # changed recently.
  if (!is.null(args$x)) {
    width_x  <- NCOL(args$x) + rnms
    height_x <- NROW(args$x) + cnms
  } else {
    width_x  <- 1
    height_x <- 1
  }

  if (length(fcol) && length(frow)) {
    if (!is.null(left)) {
      fcol <- min(fcol) - left - width_x + 1L
      frow <- min(frow)
    } else if (!is.null(right)) {
      fcol <- max(fcol) + right
      frow <- min(frow)
    } else if (!is.null(above)) {
      fcol <- min(fcol)
      frow <- min(frow) - above - height_x + 1L
    } else if (!is.null(below)) {
      fcol <- min(fcol)
      frow <- max(frow) + below
    } else {
      fcol <- max(fcol)
      frow <- max(frow)
    }
  }

  # max() on empty returns -Inf, which is caught by < 1
  if (length(fcol) == 0 || fcol < 1) {
    if (length(fcol) > 0) {
      warning("columns cannot be left of column A (integer position 1). resetting", call. = FALSE)
    }
    fcol <- 1L
  }

  if (length(frow) == 0 || frow < 1) {
    if (length(frow) > 0) {
      warning("rows cannot be above of row 1 (integer position 1). resetting", call. = FALSE)
    }
    frow <- 1L
  }

  # After this point, all unnamed problems are solved ;)
  x <- args$x
  if (!is.null(select) && is.null(args$x)) {
    stop("Can't supply `select` when `x` is absent.")
  }

  # little helper that streamlines which inputs cannot be
  select <- determine_select_valid(args = args, select = select)

  # TODO: we warn, but fail not? Shouldn't we fail?
  check_wb_dims_args(args, select = select)
  if (!is.null(rows_arg)) rows_arg <- as.integer(rows_arg)
  if (!is.null(frow))     frow     <- as.integer(frow)

  assert_class(rows_arg, class = "integer", arg_nm = "rows", or_null = TRUE)
  # Checking cols (if it is a column name)
  cols_arg <- args$cols
  x_has_named_dims <- inherits(x, "data.frame") || inherits(x, "matrix")

  if (!is.null(x)) {
    x <- as.data.frame(x, stringsAsFactors = FALSE)
  }

  cnam      <- isTRUE(args$col_names)
  col_names <- if (!is.null(args$col_names)) args$col_names else x_has_named_dims
  row_names <- if (!is.null(args$row_names)) args$row_names else FALSE

  if (cnam && !x_has_named_dims) {
    stop("Can't supply `col_names` when `x` is a vector.\n", "Transform `x` to a data.frame")
  }

  assert_class(col_names, "logical")
  assert_class(row_names, "logical")

  # special condition, can be cell reference
  if (is.null(x) && is.character(cols_arg)) {
    cols_arg <- col2int(cols_arg)
  }

  # get base dimensions
  cols_all <- if (!is.null(x)) seq_len(NCOL(x)) else cols_arg
  rows_all <- if (!is.null(x)) seq_len(NROW(x)) else rows_arg

  # get selections in this base
  cols_sel <- if (is.null(cols_arg)) cols_all else cols_arg
  rows_sel <- if (is.null(rows_arg)) rows_all else rows_arg

  # if cols is a column name from x
  if (!is.null(x) && is.character(cols_sel) && any(cols_sel %in% names(x))) {
    names(cols_all) <- names(x)
    cols_sel <- match(cols_sel, names(cols_all))
  }

  # reduce to required length
  # intersect() is faster than %in% for long vectors
  col_span <- if (is.null(cols_sel)) cols_all else intersect(cols_all, cols_sel)
  row_span <- if (is.null(rows_sel)) rows_all else intersect(rows_all, rows_sel)

  # if required add column name and row name
  if (col_names) row_span <- c(max(min(row_span, 1), 1L), row_span + 1L)
  if (row_names) col_span <- c(max(min(col_span, 1), 1L), col_span + 1L)

  # data has no column name
  if (!is.null(x) && select == "data") {
    if (col_names) row_span <- row_span[-1]
    if (row_names) col_span <- col_span[-1]
  }

  # column name has only the column name
  if (!is.null(x) && select == "col_names") {
    if (col_names) row_span <- row_span[1] # else 0?
    if (row_names) col_span <- col_span[-1]
  }

  if (!is.null(x) && select == "row_names") {
    if (col_names) row_span <- row_span[-1]
    if (row_names) col_span <- col_span[1] # else 0?
  }

  # This can happen if only from_col or from_row is supplied
  if (length(row_span) == 0) row_span <- 1L
  if (length(col_span) == 0) col_span <- 1L

  # frow & fcol start at 1
  row_span <- row_span + (frow - 1L)
  col_span <- col_span + (fcol - 1L)

  # precompute non-consecutive checks (avoid repeated diff())
  rows_nonconsec <- length(row_span) > 1L && any(abs(diff(row_span)) != 1L)
  cols_nonconsec <- length(col_span) > 1L && any(abs(diff(col_span)) != 1L)

  # return single cells (A1 or A1,B1)
  if ((length(row_span) == 1 || rows_nonconsec) &&
      (length(col_span) == 1 || cols_nonconsec)) {

    # A1
    row_start <- row_span
    col_start <- col_span

    dims <- NULL

    if (rows_nonconsec) {
      for (row_start in row_span) {
        cdims <- NULL
        if (cols_nonconsec) {
          for (col_start in col_span) {
            tmp  <- rowcol_to_dim(row_start, col_start, args$fix)
            cdims <- c(cdims, tmp)
          }
        } else {
          cdims <- rowcol_to_dim(row_start, col_span, args$fix)
        }
        dims <- c(dims, cdims)
      }
    } else {
      if (cols_nonconsec) {
        for (col_start in col_span) {
          tmp  <- rowcol_to_dim(row_span, col_start, args$fix)
          dims <- c(dims, tmp)
        }
      } else {
        dims <- rowcol_to_dim(row_span, col_span, args$fix)
      }
    }

  } else { # return range "A1:A7" or "A1:A7,B1:B7"

    dims <- NULL
    if (rows_nonconsec) {
      for (row_start in row_span) {
        cdims <- NULL
        if (cols_nonconsec) {
          for (col_start in col_span) {
            tmp  <- rowcol_to_dims(row_start, col_start, fix = args$fix)
            cdims <- c(cdims, tmp)
          }
        } else {
          cdims <- rowcol_to_dims(row_start, col_span, fix = args$fix)
        }
        dims <- c(dims, cdims)
      }
    } else {
      if (cols_nonconsec) {
        for (col_start in col_span) {
          tmp  <- rowcol_to_dims(row_span, col_start, fix = args$fix)
          dims <- c(dims, tmp)
        }
      } else {
        dims <- rowcol_to_dims(row_span, col_span, fix = args$fix)
      }
    }
  }

  # final check if any column or row exceeds the valid ranges
  validate_dims(dims)

  paste0(dims, collapse = ",")
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
  unlist(relship[c("Id")], use.names = FALSE)
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

# `fmt_txt()` ------------------------------------------------------------------
#' format strings independent of the cell style.
#'
#' @details
#' The result is an xml string. It is possible to paste multiple `fmt_txt()`
#' strings together to create a string with differing styles. It is possible to
#' supply different underline styles to `underline`.
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
# #' @param x a string or part of a string
#' @param bold,italic,underline,strike logical defaulting to `FALSE`
#' @param size the font size
#' @param color a [wb_color()] for the font
#' @param font the font name
#' @param charset integer value from the table below
#' @param vert_align any of `baseline`, `superscript`, or `subscript`
#' @param family a font family
#' @param outline,shadow,condense,extend logical defaulting to `NULL`
#' @param ... additional arguments
#' @seealso [create_font()]
#' @examples
#' fmt_txt("bar", underline = TRUE)
#' @export
fmt_txt <- function(
    x,
    bold       = FALSE,
    italic     = FALSE,
    underline  = FALSE,
    strike     = FALSE,
    size       = NULL,
    color      = NULL,
    font       = NULL,
    charset    = NULL,
    outline    = NULL,
    vert_align = NULL,
    family     = NULL,
    shadow     = NULL,
    condense   = NULL,
    extend     = NULL,
    ...
) {

  # TODO: align argument order with `create_font()`

  standardize(...)

  # CT_RPrElt
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
  xml_fmly  <- NULL
  xml_sdw   <- NULL
  xml_cdns  <- NULL
  xml_extnd <- NULL

  # For boolean elements like bold, italic, and strike: the default is a node
  # without any value. Like <b/>, not <b val="1"/>. If FALSE, we do not add a
  # node. That's what the isTRUE() and !isFALSE() checks make sure.

  if (length(bold)) {
    assert_xml_bool(bold)
    if (isTRUE(bold)) bold <- ""
    if (!isFALSE(bold))
      xml_b <-  xml_node_create("b", xml_attributes = c(val = as_xml_attr(bold)))
  }
  if (length(italic)) {
    assert_xml_bool(italic)
    if (isTRUE(italic)) italic <- ""
    if (!isFALSE(italic))
      xml_i <-  xml_node_create("i", xml_attributes = c(val = as_xml_attr(italic)))
  }
  if (length(underline)) {
    if (is.character(underline)) {
      valid_underlines <- c("single", "double", "singleAccounting", "doubleAccounting", "none")
      match.arg_wrapper(underline, valid_underlines, fn_name = "fmt_txt")
    } else {
      assert_xml_bool(strike)
      underline <- as.logical(underline) # if a numeric is passed
      if (underline) underline <- ""
    }
    if (!isFALSE(underline))
      xml_u <-  xml_node_create("u", xml_attributes = c(val = as_xml_attr(underline)))
  }
  if (length(strike)) {
    assert_xml_bool(strike)
    if (isTRUE(strike)) strike <- ""
    if (!isFALSE(strike))
      xml_strk <- xml_node_create("strike", xml_attributes = c(val = as_xml_attr(strike)))
  }
  if (length(size)) {
    xml_sz <- xml_node_create("sz", xml_attributes = c(val = as_xml_attr(size)))
  }
  if (length(color)) {
    assert_color(color)
    xml_color <- xml_node_create("color", xml_attributes = color)
  }
  if (length(font)) {
    xml_font <- xml_node_create("rFont", xml_attributes = c(val = as_xml_attr(font)))
  }
  if (length(charset)) {
    xml_chrst <- xml_node_create("charset", xml_attributes = c(val = as_xml_attr(charset)))
  }
  if (length(outline)) {
    assert_xml_bool(outline)
    if (isTRUE(outline)) outline <- ""
    xml_otln <- xml_node_create("outline", xml_attributes = c(val = as_xml_attr(outline)))
  }
  if (length(vert_align)) {
    valid_aligns <- c("baseline", "superscript", "subscript")
    match.arg_wrapper(vert_align, valid_aligns, fn_name = "fmt_txt")
    xml_vrtln <- xml_node_create("vertAlign", xml_attributes = c(val = as_xml_attr(vert_align)))
  }
  if (length(family)) {
    if (!family %in% as.character(0:14))
      stop("family needs to be in the range of 0 to 14", call. = FALSE)
    xml_fmly <- xml_node_create("family", xml_attributes = c(val = as_xml_attr(family)))
  }
  if (length(shadow)) {
    assert_xml_bool(shadow)
    if (isTRUE(shadow)) shadow <- ""
    xml_sdw <- xml_node_create("shadow", xml_attributes = c(val = as_xml_attr(shadow)))
  }
  if (length(condense)) {
    assert_xml_bool(condense)
    if (isTRUE(condense)) condense <- ""
    xml_cdns <- xml_node_create("condense", xml_attributes = c(val = as_xml_attr(condense)))
  }
  if (length(extend)) {
    assert_xml_bool(extend)
    if (isTRUE(extend)) extend <- ""
    xml_extnd <- xml_node_create("extend", xml_attributes = c(val = as_xml_attr(extend)))
  }

  xml_t_attr <- if (grepl("(^\\s+)|(\\s+$)", x)) c("xml:space" = "preserve") else NULL
  xml_t <- xml_node_create("t", xml_children = replace_legal_chars(x), xml_attributes = xml_t_attr)

  rPr_clds <- NULL
  # ommit all run properties, if x is "", this causes broken output in certain
  # spreadsheet applications
  if (nchar(x)) {
    rPr_clds <- c(
      xml_b,
      xml_i,
      xml_u,
      xml_strk,
      xml_sz,
      xml_color,
      xml_font,
      xml_chrst,
      xml_otln,
      xml_vrtln,
      xml_fmly,
      xml_sdw,
      xml_cdns,
      xml_extnd
    )
  }

  xml_rpr <- xml_node_create(
    "rPr",
    xml_children = rPr_clds
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
#' @details
#' You can join additional objects into fmt_txt() objects using "+".
#' Though be aware that `fmt_txt("sum:") + (2 + 2)` is different to `fmt_txt("sum:") + 2 + 2`.
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
#' @export
print.fmt_txt <- function(x, ...) {
  message("fmt_txt string: ")
  print(as.character(x), ...)
}

#' helper to check if a string looks like a cell
#' @param x a string
#' @keywords internal
is_dims  <- function(x) {
  grepl("^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$", x)
}

#' the function to find the cell
#' this function is used with `show_formula` in [wb_to_df()]
#' @param string the formula
#' @family internal
#' @noRd
find_a1_notation <- function(string) {
  pattern <- "\\$?[A-Z]\\$?[0-9]+(:\\$?[A-Z]\\$?[0-9]+)?"
  stringi::stri_extract_all_regex(string, pattern)
}

#' the function to replace the next cell
#' this function is used with `show_formula` in [wb_to_df()]
#' @param cell the cell from a shared formula find_a1_notation()
#' @param cols,rows an integer where to move the cell
#' @family internal
#' @noRd
next_cell <- function(cell, cols = 0L, rows = 0L) {

  z <- vector("character", length(cell))
  for (i in seq_along(cell)) {
    # Match individual cells and ranges
    match <- stringi::stri_match_first_regex(cell[[i]], "^(\\$?)([A-Z]+)(\\$?)(\\d+)(:(\\$?)([A-Z]+)(\\$?)(\\d+))?$")

    # if shared formula contains no A1 reference
    if (is.na(match[1, 1])) return(NA_character_)

    # Extract parts of the cell
    fixed_col1 <- match[2]
    col1 <- match[3]
    fixed_row1 <- match[4]
    row1 <- as.numeric(match[5])

    fixed_col2 <- match[7]
    col2 <- match[8]
    fixed_row2 <- match[9]
    row2 <- as.numeric(match[10])

    if (is.na(col2)) {

      # Handle individual cell
      next_col <- if (fixed_col1 == "") int2col(col2int(col1) + cols) else col1
      next_row <- if (fixed_row1 == "") row1 + rows else row1
      z[i] <- paste0(fixed_col1, next_col, fixed_row1, next_row)

    } else {

      # Handle range
      next_col1 <- if (fixed_col1 == "") int2col(col2int(col1) + cols) else col1
      next_row1 <- if (fixed_row1 == "") row1 + rows else row1
      next_col2 <- if (fixed_col2 == "") int2col(col2int(col2) + cols) else col2
      next_row2 <- if (fixed_col2 == "") row2 + rows else row2
      z[i] <- paste0(fixed_col1, next_col1, fixed_row1, next_row1, ":", fixed_col2, next_col2, fixed_row2, next_row2)

    }
  }

  z
}

#' the replacement function for shared formulas
#' this function is used with `show_formula` in [wb_to_df()]
#' @param string the formula
#' @param replacements the replacements, from next_cell()
#' @family internal
#' @noRd
replace_a1_notation <- function(strings, replacements) {

  # create sprintf-able strings
  strings <- stringi::stri_replace_all_regex(
    strings,
    "\\$?[A-Z]\\$?[0-9]+(:\\$?[A-Z]\\$?[0-9]+)?",
    "%s"
  )

  # insert replacements into string
  repl_fun <- function(str, y) {
    if (!anyNA(y)) # else keep str as is
      str <- do.call(sprintf, c(str, as.list(y)))
    str
  }

  z <- vector("character", length(strings))
  for (i in seq_along(strings)) {
    z[i] <- repl_fun(strings[[i]], replacements[[i]])
  }

  z
}

# extend shared formula into all formula cells
carry_forward <- function(x) {
  rep(x[1], length(x))
}

# calculate difference for each shared formula to the origin
calc_distance <- function(x) {
  x - x[1]
}

# safer rbind
rbind2 <- function(df1, df2) {
  if (is.null(df1)) return(df2)
  nms <- unique(c(names(df1), names(df2)))
  df1[setdiff(nms, names(df1))] <- ""
  df2[setdiff(nms, names(df2))] <- ""
  rbind(df1[nms], df2[nms])
}

# ave function to avoid a dependency on stats. if we ever rely on stats,
# this can be replaced by stats::ave
ave2 <- function(x, y, FUN) {
  g <- as.factor(y)
  split(x, g) <- lapply(split(x, g), FUN)
  x
}

# file_ext function to avoid a depencency on tools. if we ever rely on tools,
# this can be replaced by tools::file_ext
file_ext2 <- function(filepath) {
  sub(".*\\.", "", basename2(filepath))
}

if (getRversion() < "4.0.0") {
  deparse1 <- function(expr, collapse = " ") {
    paste(deparse(expr), collapse = collapse)
  }
}

## the R solution
utilszip <- function(zip_path, source_dir, compression_level = 9) {
  if (!is.null(getOption("openxlsx2.debug"))) message("utils::zip")
  original_wd <- getwd()
  on.exit(setwd(original_wd), add = TRUE)
  abs_zip_path <- normalizePath(zip_path, mustWork = FALSE)
  setwd(source_dir)
  zip_flags <- paste0("-r", compression_level, "Xq")
  opt_flags <- getOption("openxlsx2.zip_flags")
  if (!is.null(opt_flags)) zip_flags <- opt_flags
  source_files <- list.files(source_dir, full.names = FALSE)
  utils::zip(zipfile = abs_zip_path, files = source_files, flags = zip_flags)
}

## bsdtar seems to be the best fallback solution
# Windows ships bsdtar since Windows 10, Apple has an older bsdtar
# On Linux it should be available if libarchive is installed
bsdtar <- function(zip_path, source_dir, compression_level = 9) {
  if (!is.null(getOption("openxlsx2.debug"))) message("bsdtar")
  original_wd <- getwd()
  on.exit(setwd(original_wd), add = TRUE)
  abs_zip_path <- normalizePath(zip_path, mustWork = FALSE)
  setwd(source_dir)
  tar_args <- paste0(
    "-a -c -f ", abs_zip_path, " --options zip:compression-level=",
    compression_level, " --no-acls --no-xattrs --no-fflags --format=zip")
  opt_flags <- getOption("openxlsx2.zip_flags")
  if (!is.null(opt_flags)) tar_args <- opt_flags
  tar_args <- c(tar_args, list.files(source_dir, full.names = FALSE))
  command <- if (Sys.info()[["sysname"]] == "Windows")
      Sys.which("tar")
    else
      Sys.which("bsdtar")
  system2(command, args = tar_args)
}

## as an alternative solution
p7zip <- function(zip_path, source_dir, compression_level = 9) {
  if (!is.null(getOption("openxlsx2.debug"))) message("7zip")
  original_wd <- getwd()
  on.exit(setwd(original_wd), add = TRUE)
  abs_zip_path <- normalizePath(zip_path, mustWork = FALSE)
  setwd(source_dir)
  # default: quiet, suppress all output similar to -r9Xq
  zip_args <- sprintf("a -tzip -r -bb0 -bso0 -sns- -mx=%s", compression_level)
  zip_args <- getOption("openxlsx2.zip_flags", default = zip_args)
  zip_args <- c(zip_args, abs_zip_path, list.files(source_dir, full.names = FALSE))
  system2(Sys.getenv("R_ZIPCMD"), args = zip_args)
}

## the previous solution
zipzip <- function(zip_path, source_dir, compression_level = 9) {
  if (!is.null(getOption("openxlsx2.debug"))) message("zip::zip")
  suppressMessages(requireNamespace("zip"))
  zip::zip(
    zipfile = zip_path,
    files = list.files(source_dir, full.names = FALSE),
    recurse = TRUE,
    compression_level = compression_level,
    include_directories = FALSE,
    # change the working directory for this
    root = source_dir,
    # change default to match historical zipr
    mode = "cherry-pick"
  )
}

## previously we were relying on the zip package for R. There is nothing wrong
# with the package, but it is another dependency that could be avoided using
# reliable system tools. Unfortunately the utils::zip() function does not ship
# a zip tool and now we are in a mess where we have to check if a zip tool is
# available. And there are several cases where it is not
zip_output <- function(zip_path, source_dir, compression_level = 9) {

  # on Windows we might have an Rtools folder somewhere with a zip.exe ...
  if (.Platform$OS.type == "windows" && Sys.getenv("R_ZIPCMD") == "") {
    # Windows 10 pre 2017, Windows 7 or older
    env <- Sys.getenv()
    env <- env[grepl("^RTOOLS[A-Z0-9_]+_HOME$", names(env))]
    maybe_zip <- file.path(env, "usr", "bin", "zip.exe")
    if (any(zip_here <- file.exists(maybe_zip)))
      Sys.setenv("R_ZIPCMD" = maybe_zip[zip_here][1])
    on.exit(Sys.setenv("R_ZIPCMD" = ""), add = TRUE)
  }

  if (!isTRUE(getOption("openxlsx2.no_utils_zip"))) {
    if (Sys.which("zip") != "" || Sys.getenv("R_ZIPCMD") != "") {
      res <- utilszip(
        zip_path = zip_path,
        source_dir = source_dir,
        compression_level = compression_level
      )
      return(res)
    }
  }

  if (!isTRUE(getOption("openxlsx2.no_bsdtar"))) {
    if (grepl("WINDOWS", Sys.which("tar")) || Sys.which("bsdtar") != "") {
      res <- bsdtar(
        zip_path = zip_path,
        source_dir = source_dir,
        compression_level = compression_level
      )
      return(res)
    }
  }

  if (grepl("7z", Sys.getenv("R_ZIPCMD"))) {
    res <- p7zip(
      zip_path = zip_path,
      source_dir = source_dir,
      compression_level = compression_level
    )
    return(res)
  }

  res <- zipzip(
    zip_path = zip_path,
    source_dir = source_dir,
    compression_level = compression_level
  )
  res
}
