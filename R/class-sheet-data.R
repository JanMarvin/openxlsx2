

SheetData <- setRefClass(
  "SheetData",
  # TODO should all fields be initiated?  Why would any not be?
  fields = c(
    rows       = "integer",
    cols       = "integer",
    s          = "ANY", # not initialized
    r          = "ANY", # not initialized
    t          = "ANY",
    v          = "ANY",
    f          = "ANY",
    row_attr   = "ANY",
    cc         = "ANY", # not initialized
    cc_out     = "ANY", # not initialized
    style_id   = "ANY",
    data_count = "integer",
    n_elements = "integer"
  ),

  methods = list(
    # TODO should SheetData$new() have params?
    initialize = function() {
      .self$rows <- integer()
      .self$cols <- integer()

      .self$t <- integer()
      .self$v <- character()
      .self$f <- character()

      .self$style_id <- character()

      .self$data_count <- 0L
      .self$n_elements <- 0L

      invisible(.self)
    },

    delete = function(rows_in, cols_in, grid_expand) {
      cols_in <- convertFromExcelRef(cols_in)
      rows_in <- as.integer(rows_in)

      ## rows and cols need to be the same length
      if (grid_expand) {
        n <- length(rows_in)
        rows_in <- rep.int(rows_in, times = length(cols_in))
        cols_in <- rep(cols_in, each = n)
      }

      if (length(rows_in) != length(cols_in)) {
        stop("Length of rows and cols must be equal.")
      }

      inds <- which(paste(rows, cols, sep = ",") %in% paste(rows_in, cols_in, sep = ","))

      if (length(inds) > 0) { ## writing over existing data

        .self$rows <- rows[-inds]
        .self$cols <- cols[-inds]
        .self$t <- t[-inds]
        .self$v <- v[-inds]
        .self$f <- f[-inds]

        .self$n_elements <- as.integer(length(rows))

        if (n_elements == 0) {
          .self$data_count <- 0L
        }
      }

      invisible(.self)
    },

    write = function(rows_in, cols_in, t_in, v_in, f_in, any_functions = TRUE) {
      if (length(rows_in) == 0 | length(cols_in) == 0) {
        return(invisible(.self))
      }


      possible_overlap <- FALSE
      if (n_elements > 0) {
        possible_overlap <- (min(cols_in, na.rm = TRUE) <= max(cols, na.rm = TRUE)) &
          (max(cols_in, na.rm = TRUE) >= min(cols, na.rm = TRUE)) &
          (min(rows_in, na.rm = TRUE) <= max(rows, na.rm = TRUE)) &
          (max(rows_in, na.rm = TRUE) >= min(rows, na.rm = TRUE))
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
      inds <- integer()
      if (possible_overlap) {
        inds <- which(paste(rows, cols, sep = ",") %in% paste(rows_in, cols_in, sep = ","))
      }

      if (length(inds) > 0) {
        .self$rows <- c(.self$rows[-inds], rows_in)
        .self$cols <- c(.self$cols[-inds], cols_in)
        .self$t <- c(.self$t[-inds], t_in)
        .self$v <- c(.self$v[-inds], v_in)
        .self$f <- c(.self$f[-inds], f_in)
      } else {
        .self$rows <- c(.self$rows, rows_in)
        .self$cols <- c(.self$cols, cols_in)
        .self$t <- c(.self$t, t_in)
        .self$v <- c(.self$v, v_in)
        .self$f <- c(.self$f, f_in)
      }

      .self$n_elements <- as.integer(length(rows))
      .self$data_count <- .self$data_count + 1L
      invisible(.self)
    }
  )
)

new_sheet_data <- function() {
  SheetData$new()
}


# helpers -----------------------------------------------------------------

# Consider making some helpers for the cc stuff.

empty_sheet_data_cc <- function() {
  # make make this a specific class/object?
  stop("this isn't in use")
  data.frame(
    row_r = character(),
    c_r   = character(),
    c_s   = character(),
    c_t   = character(),
    v     = character(),
    f     = character(),
    f_t   = character(),
    f_ref = character(),
    f_si  = character(),
    is    = character(),
    typ   = character(),
    r     = character(),
    stringsAsFactors = FALSE
  )
}


empty_row_attr <- function(n = 0) {
  # make make this a specific class/object?

  row_attr_nams <- c("collapsed", "customFormat", "customHeight",
                     "x14ac:dyDescent", "ht", "hidden", "outlineLevel",
                     "r", "ph", "spans", "s", "thickBot", "thickTop")

  z <- data.frame(
    matrix("", nrow = n, ncol = length(row_attr_nams)),
    stringsAsFactors = FALSE
  )
  names(z) <- row_attr_nams

  z
}
