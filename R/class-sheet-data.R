
#' R6 class for a Workbook Hyperlink
#'
#' A hyperlink
#'
#' @export
wbSheetData <- R6::R6Class(
  "wbSheetData",
  public = list(
    # TODO which fields should be moved to private?

    #' @field rows rows
    rows = 0L,

    #' @field cols cols
    cols = 0L,

    # TODO better descriptions for s, r, t, v, cc, cc_out

    #' @field s s
    s = NULL, # not initialized

    #' @field r r
    r = NULL, # not initialized

    #' @field t t
    t = integer(),

    #' @field v v
    v = character(),

    #' @field f f
    f = character(),

    #' @field row_attr row_attr
    row_attr = NULL,

    #' @field cc cc
    cc = NULL,

    #' @field cc_out cc_out
    cc_out = NULL,

    #' @field style_id style_id
    style_id = character(),

    #' @field data_count data_count

    data_count = 0L,

    #' @field n_elements n_elements
    n_elements = 0L,

    #' @description
    #' Creates a new `wbSheetData` object
    #' @return a `wbSheetData` object
    initialize = function() {
      self$rows <- 0L
      self$cols <- 0L

      self$t <- integer()
      self$v <- character()
      self$f <- character()

      self$style_id <- character()

      self$data_count <- 0L
      self$n_elements <- 0L

      self
    },

    #' @description
    #' Delete data
    #'
    #' @param rows_in,cols_in The row and columns to delete
    #' @param grid_expand If `TRUE` will expand the selection of `rows_in`, and
    #'   `cols_in` to the full set, not just the exact positions (similar to
    #'    using `[expand.grid()]`)
    #' @return The `wbSheetData` object, invisibly
    delete = function(rows_in, cols_in, grid_expand) {
      cols_in <- convertFromExcelRef(cols_in)
      rows_in <- as.integer(rows_in)

      ## rows and cols need to be the same length
      # TODO maybe make grid_expand into a function? (not expand.grid)? Used in
      # several locations
      if (grid_expand) {
        n <- length(rows_in)
        rows_in <- rep.int(rows_in, times = length(cols_in))
        cols_in <- rep(cols_in, each = n)
      }

      if (length(rows_in) != length(cols_in)) {
        stop("Length of rows and cols must be equal.")
      }

      inds <- paste(self$rows, self$cols, sep = ",") %in% paste(rows_in, cols_in, sep = ",")
      inds <- which(inds)

      if (length(inds)) { ## writing over existing data

        self$rows <- self$rows[-inds]
        self$cols <- self$cols[-inds]
        self$t <- self$t[-inds]
        self$v <- self$v[-inds]
        self$f <- self$f[-inds]

        self$n_elements <- as.integer(length(self$rows))

        if (self$n_elements == 0) {
          self$data_count <- 0L
        }
      }

      invisible(self)
    },

    # TODO what are f_in, v_in, f_in, and any_functions?

    #' @description
    #' Write data
    #'
    #' @param rows_in,cols_in The row and columns to delete
    #' @param t_in,v_in,f_in specifications ???
    #' @param any_functions Logical ???
    #' @return The `wbSheetData` object, invisibly
    write = function(rows_in, cols_in, t_in, v_in, f_in, any_functions = TRUE) {
      if (length(rows_in) == 0 | length(cols_in) == 0) {
        return(invisible(self))
      }

      possible_overlap <- FALSE
      if (n_elements > 0) {
        possible_overlap <-
          (min(cols_in, na.rm = TRUE) <= max(self$cols, na.rm = TRUE)) &
          (max(cols_in, na.rm = TRUE) >= min(self$cols, na.rm = TRUE)) &
          (min(rows_in, na.rm = TRUE) <= max(self$rows, na.rm = TRUE)) &
          (max(rows_in, na.rm = TRUE) >= min(self$rows, na.rm = TRUE))
      }

      n <- length(cols_in)
      cols_in <- rep.int(cols_in, times = length(rows_in))
      rows_in <- rep(rows_in, each = n)

      if (any_functions) {
        if (any(!is.na(f_in))) {
          v_in[!is.na(f_in)] <- NA_character_
          t_in[!is.na(f_in)] <- 3L ## "str"
        }
      }

      # TODO change inds to ! x %in% y; don't need length() check
      if (possible_overlap) {
        # inds <- paste(self$rows, self$cols, sep = ",") %in% paste(rows_in, cols_in, sep = ",")
        inds <- which(self$rows %in% rows_in & self$cols %in% cols_in)
      } else {
        inds <- integer()
      }

      if (length(inds)) {
        self$rows <- c(self$rows[-inds], rows_in)
        self$cols <- c(self$cols[-inds], cols_in)
        self$t    <- c(self$t[-inds], t_in)
        self$v    <- c(self$v[-inds], v_in)
        self$f    <- c(self$f[-inds], f_in)
      } else {
        self$rows <- c(self$rows, rows_in)
        self$cols <- c(self$cols, cols_in)
        self$t    <- c(self$t, t_in)
        self$v    <- c(self$v, v_in)
        self$f    <- c(self$f, f_in)
      }

      self$n_elements <- as.integer(length(self$rows))
      self$data_count <- self$data_count + 1L
      invisible(self)
    }
  )
)

#' @rdname wbSheetData
#' @export
wb_sheet_data <- function() {
  wbSheetData$new()
}


# helpers -----------------------------------------------------------------

# Consider making some helpers for the cc stuff.

empty_sheet_data_cc <- function(n = 0) {
  value <- rep.int(NA_character_, n)
  data.frame(
    row_r = value,
    c_r   = value,
    c_s   = value,
    c_t   = value,
    v     = value,
    f     = value,
    f_t   = value,
    f_ref = value,
    f_si  = value,
    is    = value,
    typ   = value,
    r     = value,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
}


empty_row_attr <- function(n = 0) {
  value <- character(n)
  data.frame(
    collapsed         = value,
    customFormat      = value,
    customHeight      = value,
    `x14ac:dyDescent` = value,
    ht                = value,
    hidden            = value,
    outlineLevel      = value,
    r                 = value,
    ph                = value,
    spans             = value,
    s                 = value,
    thickBot          = value,
    thickTop          = value,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
}
