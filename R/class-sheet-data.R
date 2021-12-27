

SheetData <- setRefClass("SheetData",
  fields = c(
    "rows" = "integer",
    "cols" = "integer",
    "s" = "ANY",
    "r" = "ANY",
    "t" = "ANY",
    "v" = "ANY",
    "f" = "ANY",
    "row_attr" = "ANY",
    "cc" = "ANY",
    "cc_out" = "ANY",
    "style_id" = "ANY",
    "data_count" = "integer",
    "n_elements" = "integer"
  ),

  methods = list(
    initialize = function() {
      .self$rows <- integer(0)
      .self$cols <- integer(0)

      .self$t <- integer(0)
      .self$v <- character(0)
      .self$f <- character(0)

      .self$style_id <- character(0)

      .self$data_count <- 0L
      .self$n_elements <- 0L
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
    },

    write = function(rows_in, cols_in, t_in, v_in, f_in, any_functions = TRUE) {
      if (length(rows_in) == 0 | length(cols_in) == 0) {
        return(invisible(0))
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
          v_in[!is.na(f_in)] <- as.character(NA)
          t_in[!is.na(f_in)] <- 3L ## "str"
        }
      }

      inds <- integer(0)
      if (possible_overlap) {
        inds <- which(paste(rows, cols, sep = ",") %in% paste(rows_in, cols_in, sep = ","))
      }

      if (length(inds) > 0) {
        .self$rows <- c(rows[-inds], rows_in)
        .self$cols <- c(cols[-inds], cols_in)
        .self$t <- c(t[-inds], t_in)
        .self$v <- c(v[-inds], v_in)
        .self$f <- c(f[-inds], f_in)
      } else {
        .self$rows <- c(rows, rows_in)
        .self$cols <- c(cols, cols_in)
        .self$t <- c(t, t_in)
        .self$v <- c(v, v_in)
        .self$f <- c(f, f_in)
      }

      .self$n_elements <- as.integer(length(rows))
      .self$data_count <- data_count + 1L
    }
  )
)

