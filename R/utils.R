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
#' wb_dims(1:10, 1:10)
#'
#' @name dims_helper
#' @seealso [wb_dims()]
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
check_wb_dims_args <- function(args, select = NULL) {
  select <- match.arg(select, c("x", "data", "col_names", "row_names"))

  cond_acceptable_len_1 <- !is.null(args$from_row) || !is.null(args$from_col) || !is.null(args$x)
  nams <- names(args) %||% rep("", length(args))
  all_args_unnamed <- all(!nzchar(nams))

  if (length(args) == 1 && !cond_acceptable_len_1) {
    # Providing a single argument acceptable is only  `x`
    sentence_unnamed <- ifelse(all_args_unnamed, " unnamed ", " ")
    stop(
      "Supplying a single", sentence_unnamed, "argument to `wb_dims()` is not supported. \n",
      "Use any of `x`, `from_row` `from_col`. You can also use `rows` and `cols`, or `dims = NULL`",
      call. = FALSE
    )
  }
  cnam_null <- is.null(args$col_names)
  rnam_null <- is.null(args$row_names)
  if (is.character(args$rows) || is.character(args$from_row)) {
    warning("`rows` and `from_rows` in `wb_dims()` should not be a character. Please supply an integer vector.", call. = FALSE)
  }

  if (is.null(args$x)) {
    if (!cnam_null || !rnam_null) {
      stop("In `wb_dims()`, `row_names`, and `col_names` should only be used if `x` is present.", call. = FALSE)
    }
  }

  x_has_colnames <- !is.null(colnames(args$x))

  if (x_has_colnames && !is.null(args$rows) && is.character(args$rows)) {
    # Not checking whether it's a row name, not supported.
    is_rows_a_colname <- args$row %in% colnames(args$x)

    if (any(is_rows_a_colname)) {
      stop(
        "`rows` is the incorrect argument in this case\n",
        "Use `cols` instead. Subsetting rows by name is not supported.",
        call. = FALSE
      )
    }
  }
  invisible(NULL)
}

# it is a wrapper around base::match.arg(), but it doesn't allow partial matching.
# It also provides a more informative error message in case it fails.
match.arg_wrapper <- function(arg, choices, several.ok = FALSE, fn_name = NULL) {
  # Check valid argument names
  # partial matching accepted
  fn_name <- fn_name %||% "fn_name"

  if (!several.ok) {
    if (length(arg) != 1) {
      stop(
        "Must provide a single argument found in ", fn_name, ": ",
        invalid_arg_nams, "\n", "Use one of ", valid_arg_nams,
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
  valid_cases_choices <- names(valid_cases)
  match.arg_wrapper(select, choices = valid_cases_choices, fn_name = "wb_dims", several.ok = FALSE)

  if (isFALSE(valid_cases[[select]])) {
    stop(
      "You provided a bad value to `select` in `wb_dims()`.\n ",
      "Please review. see `?wb_dims`.",
      call. = FALSE
    )
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
#' 3. provide the `dims` of column names `wb_dims(x = mtcars, select = "col_names)`
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
#' * `from_row` / `from_col` the starting position of `x`
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
#' @param ... construct `dims` arguments, from rows/cols vectors or objects that
#'   can be coerced to data frame. `x`, `rows`, `cols`, `from_row`, `from_col`,
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
#'
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
wb_dims <- function(..., select = NULL) {
  args <- list(...)
  len <- length(args)

  if (len == 0 || (len == 1 && is.null(args[[1]]))) {
    stop("`wb_dims()` requires `rows`, `cols`, `from_row`, `from_col`, or `x`.")
    return("A1")
  }

  # nams cannot be NULL now
  nams <- names(args) %||% rep("", len)
  valid_arg_nams <- c("x", "rows", "cols", "from_row", "from_col", "row_names", "col_names")
  any_args_named <- any(nzchar(nams))
  # unused, but can be used, if we need to check if any, but not all
  # Check if valid args were provided if any argument is named.
  if (any_args_named) {
    if (any(c("start_col", "start_row") %in% nams)) {
      stop("Use `from_row` / `from_col` instead of `start_row` / `start_col`")
    }
    match.arg_wrapper(arg = nams, choices = c(valid_arg_nams, ""), several.ok = TRUE, fn_name = "`wb_dims()`")
  }
  # After this point, no need to search for invalid arguments!

  n_unnamed_args <- length(which(!nzchar(nams)))
  all_args_unnamed <- n_unnamed_args == len
  # argument dispatch / creation here.
  # All names provided, happy :)
  # Checking if valid names were provided.

  if (n_unnamed_args > 2) {
    stop("Only `rows` and `cols` can be provided unnamed. You must name all other arguments.")
  }
  if (len == 1 && all_args_unnamed) {
    stop(
      "Supplying a single unnamed argument is not handled by `wb_dims()`",
      "use `x`, `from_row` / `from_col`. You can also use `dims = NULL`"
    )
  }

  ok_if_arg1_unnamed <-
    is.atomic(args[[1]]) || any(nams %in% c("rows", "cols"))

  if (nams[1] == "" && !ok_if_arg1_unnamed) {
    stop(
      "The first argument must either be named or be a vector.",
      "Providing a single named argument must either be `from_row`, `from_col` or `x`."
    )
  }

  if (n_unnamed_args == 1 && len > 1 && !"rows" %in% nams) {
    message("Assuming the first unnamed argument to be `rows`.")
    nams[which(nams == "")[1]] <- "rows"
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == len
  }

  if (n_unnamed_args == 1 && len > 1 && "rows" %in% nams) {
    message("Assuming the first unnamed argument to be `cols`.")
    nams[which(nams == "")[1]] <- "cols"
    names(args) <- nams
    n_unnamed_args <- length(which(!nzchar(nams)))
    all_args_unnamed <- n_unnamed_args == len
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
  }

  # Just keeping this as a safeguard
  has_some_unnamed_args <- any(!nzchar(nams))
  if (has_some_unnamed_args) {
    stop("Internal error, all arguments should be named after this point.")
  }

  # After this point, all unnamed problems are solved ;)
  x <- args$x
  if (!is.null(select) && is.null(args$x)) {
    stop("`select` should only be provided with `x`.")
  }

  # little helper that streamlines which inputs cannot be
  select <- determine_select_valid(args = args, select = select)

  check_wb_dims_args(args, select = select)
  rows_arg <- args$rows
  rows_arg <- if (is.character(rows_arg)) {
    col2int(rows_arg)
  } else if (!is.null(rows_arg)) {
    as.integer(rows_arg)
  } else if (!is.null(args$x)) {
    # rows_arg <- seq_len(nrow(args$x))
    rows_arg <- NULL
  } else {
    1L
  }

  assert_class(rows_arg, class = "integer", arg_nm = "rows", or_null = TRUE)
  # Checking cols (if it is a column name)
  cols_arg <- args$cols
  x_has_named_dims <- inherits(x, "data.frame") || inherits(x, "matrix")
  x_has_colnames <- !is.null(colnames(x))
  if (!is.null(x)) {
    x <- as.data.frame(x)
  }

  cnam_null <- is.null(args$col_names)
  col_names <- args$col_names %||% x_has_named_dims

  if (!cnam_null && !x_has_named_dims) {
    stop("Supplying `col_names` when `x` is a vector is not supported.")
  }

  row_names <- args$row_names %||% FALSE
  assert_class(col_names, "logical")
  assert_class(row_names, "logical")

  # Find column location id if `cols` is a character and is a colname of x
  if (x_has_colnames && !is.null(cols_arg)) {
    is_cols_a_colname <- cols_arg %in% colnames(x)

    if (any(is_cols_a_colname)) {
      if (length(is_cols_a_colname) != 1) {
        stop(
          "Supplying multiple column names is not supported by the `wb_dims()` helper, ",
          "use the `cols` with a range instead of `x` column names.",
          "\n Use a single `cols` at a time with `wb_dims()`"
        )
      }
      # message("Transforming col name = '", cols_arg, "' to `cols = ", which(colnames(x) == cols_arg), "`")
      cols_arg <- which(colnames(x) == cols_arg)
    }
  }

  if (!is.null(cols_arg)) {
    cols_arg <- col2int(cols_arg)
    assert_class(cols_arg, class = "integer", arg_nm = "cols")
  } else if (!is.null(args$x)) {
    cols_arg <- NULL
  } else {
    cols_arg <- 1L # no more NULL for cols_arg and rows_arg if `x` is not supplied
  }

  if (!is.null(cols_arg) && (min(cols_arg) < 1L || (length(cols_arg) > 1 && any(diff(cols_arg) != 1)))) {
    stop("You must supply positive, consecutive values to `cols`")
  }

  if (!is.null(rows_arg) && (min(rows_arg) < 1L || (length(rows_arg) > 1 && any(diff(rows_arg) != 1)))) {
    stop("You must supply positive, consecutive values to `rows`.")
  }

  # assess from_row / from_col
  if (is.character(args$from_row)) {
    frow <- col2int(args$from_row)
  } else {
    frow <- args$from_row %||% 1L
    frow <- as.integer(frow)
  }

  # from_row is a function of col_names, from_rows and cols.
  # cols_seq should start at 1 after this
  # if from_row = 4, rows = 4:7,
  # then frow = 4 + 4 et rows = seq_len(length(rows))
  fcol <- col2int(args$from_col) %||% 1L
  # after this point, no assertion, assuming all elements to be acceptable

  # from_row / from_col = 0 only acceptable in certain cases.
  if (!all(length(fcol) == 1, length(frow) == 1, fcol >= 1, frow >= 1)) {
    stop("`from_col` / `from_row` should have length 1. and be positive.")
  }

  if (select == "col_names") {
    ncol_to_span <- ncol(x)
    nrow_to_span <- 1L
  } else if (select == "row_names") {
    ncol_to_span <- 1L
    nrow_to_span <- nrow(x) %||% 1L
  } else if (select %in% c("x", "data")) {
    if (!is.null(cols_arg)) {
      ncol_to_span <- length(cols_arg)
    } else {
      ncol_to_span <- ncol(x) %||% 1L
    }
    if (!is.null(rows_arg)) {
      nrow_to_span <- length(rows_arg)
    } else {
      nrow_to_span <- nrow(x) %||% 1L
    }

    if (select == "x") {
      nrow_to_span <- nrow_to_span + col_names
      ncol_to_span <- ncol_to_span + row_names
    }
  }

  # Setting frow / fcol correctly.
  if (select == "row_names") {
    fcol <- fcol
    frow <- frow + col_names
  } else if (select == "col_names") {
    fcol <- fcol + row_names
    frow <- frow
  } else if (select %in% c("x", "data")) {
    if (!is.null(cols_arg)) {
      if (min(cols_arg) > 1) {
        fcol <- fcol + min(cols_arg) - 1L
      }
    }

    if (!is.null(rows_arg)) {
      if (min(rows_arg) > 1) {
        frow <- frow + min(rows_arg) - 1L
      }
    }

    if (select == "data") {
      fcol <- fcol + row_names
      frow <- frow + col_names
    }
  }

  row_span <- frow + seq_len(nrow_to_span) - 1L
  col_span <- fcol + seq_len(ncol_to_span) - 1L

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
  xml_t <- xml_node_create("t", xml_children = replace_legal_chars(x), xml_attributes = xml_t_attr)

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
