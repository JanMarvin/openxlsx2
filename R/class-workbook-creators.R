workbook_modify_creators <- function(
    self,
    private,
    method = c("add", "set", "remove"),
    value
) {
  method <- match.arg(method)
  assert_class(value, "character")

  if (any(!has_chr(value))) {
    stop("all creators must contain characters without NAs", call. = FALSE)
  }

  value <- switch(
    method,
    add    = unique(c(self$creator, value)),
    set    = unique(value),
    remove = setdiff(self$creator, value)
  )

  self$creator <- value
  # core is made on initialization
  private$generate_base_core()
  invisible(self)
}
