
#' takes colour and returns color
#' @param ... ...
#' @returns void. assigns an object in the parent frame
#' @noRd
standardize_color_names <- function(..., return = FALSE) {

  # since R 4.1.0: ...names()
  args <- list(...)
  if (return) {
    args <- args[[1]]
  }
  got <- names(args)
  # can be Color or color
  got_color <- grep("colour", tolower(got))

  if (length(got_color)) {
    for (got_col in got_color) {
      color <- got[got_col]
      name_color <- stringi::stri_replace_all_fixed(color, "olour", "olor")

      if (return) {
        names(args)[got_col] <- name_color
      } else {
        # since R 3.5.0: ...elt(got_col)
        value_color <- args[[got_col]]
        assign(name_color, value_color, parent.frame())
      }
    }
  }

  if (return) args
}

#' takes camelCase and returns camel_case
#' @param ... ...
#' @returns void. assigns an object in the parent frame
#' @noRd
standardize_case_names <- function(..., return = FALSE, arguments = NULL) {

  if (is.null(arguments)) {
    arguments <- ls(envir = parent.frame())
  }

  # since R 4.1.0: ...names()
  args <- list(...)
  if (return) {
    args <- args[[1]]
  }
  got <- names(args)

  regex <- "(\\G(?!^)|\\b[a-zA-Z][a-z]*)([A-Z][a-z]*|\\d+)"

  got_camel_cases <- grep(regex, got, perl = TRUE)

  if (length(got_camel_cases)) {
    for (got_camel_case in got_camel_cases) {
      camel_case <- got[got_camel_case]
      name_camel_case <- gsub(
        pattern     = regex,
        replacement = "\\L\\1_\\2",
        x           = camel_case,
        perl        = TRUE
      )
      got[got_camel_case] <- name_camel_case
      # since R 3.5.0: ...elt(got_col)
      if (return) {
        names(args)[got_camel_case] <- name_camel_case
      } else {
        value_camel_calse <- args[[got_camel_case]]
        assign(name_camel_case, value_camel_calse, parent.frame())
      }
    }
    if (getOption("openxlsx2.soon_deprecated", default = FALSE)) {
      msg <- paste(
        "Found camelCase arguments in code.",
        "These will be deprecated in the next major release.",
        "Consider using:", paste(got[got_camel_cases], collapse = ", ")
      )
      .Deprecated(msg = msg)
    }
  }

  sel <- !got %in% arguments
  if (any(sel)) {
    warning("unused arguments (", paste(got[sel], collapse = ", "), ")")
  }

  if (return) args

}

#' takes camelCase and colour returns camel_case and color
#' @param ... ...
#' @returns void. assigns an object in the parent frame
#' @noRd
standardize <- function(..., arguments) {

  nms <- list(...)
  if (missing(arguments)) {
    arguments <- ls(envir = parent.frame())
  }

  rtns <- standardize_color_names(nms, return = TRUE)
  rtns <- standardize_case_names(rtns, return = TRUE, arguments = arguments)

  nms <- names(rtns)
  for (i in seq_along(nms)) {
    assign(nms[[i]], rtns[[i]], parent.frame())
  }
}
