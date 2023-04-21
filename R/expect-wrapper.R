#' Expect wrapper
#'
#' Internal tool:  see tests/testthat/test-workbook-wrappers.R
#'
#' @description Testing expectations for [wbWorkbook] wrappers.  These test that
#'   the method and wrappers have the same params, same default values, and
#'   return the same value.
#'
#'   Requires the `waldo` package, used within `testthat`.
#'
#' @param method The name of the [wbWorkbook] method
#' @param fun The name of the wrapper (probably in R/class-workbook-wrappers)
#' @param wb A `wbWorkbook` object
#' @param params A named list of params.  Set to `NULL` to not execute functions
#' @param ignore Names of params to ignore
#' @param ignore_attr Passed to `waldo::compare()` for the result of the
#'   executed functions
#' @param ignore_fields Ignore immediate fields by removing them from the
#'   objects for comparison
#' @param ignore_wb for pseudo wbWorkbook wrappers such as wb_load() we have
#'   to ignore wb
#' @returns Nothing, called for its side-effects
#' @rdname expect_wrapper
#' @keywords internal
#' @noRd
expect_wrapper <- function(
  method,
  fun           = paste0("wb_", method),
  wb            = wb_workbook(),
  params        = list(),
  ignore        = NULL,
  ignore_attr   = "waldo_opts",
  ignore_fields = NULL,
  ignore_wb     = FALSE
) {
  stopifnot(
    requireNamespace("waldo", quietly = TRUE),
    requireNamespace("testthat", quietly = TRUE),
    is.character(method), length(method) == 1L,
    is.character(fun),    length(fun) == 1L,
    is.list(params) || is.null(params)
  )

  method_fun <- get(method, wbWorkbook$public_methods)
  fun_fun    <- match.fun(fun)

  # these come as pairlist -- we don't really care about that
  method_forms <- as.list(formals(method_fun))
  fun_forms    <- as.list(formals(fun_fun))

  method_args <- names(method_forms)
  fun_args    <- names(fun_forms)

  # remove wb and other ignores
  ignore <- c("wb", ignore)

  # remove ignores from fun
  m <- match(ignore, fun_args, 0L)
  if (!identical(m, 0L)) {
    # remove wb from both the args and formals
    fun_args  <- fun_args[-m]
    fun_forms <- fun_forms[-m]
  }

  # remove ignores from method
  m <- match(ignore, method_args, 0L)
  if (!identical(m, 0L)) {
    # remove wb from both the args and formals
    method_args  <- method_args[-m]
    method_forms <- method_forms[-m]
  }

  method0 <- paste0("wbWorkbook$", method)

  # adjustments for when wrappers don't have any args
  if (!length(method_args)) {
    method_args <- character()
  }

  if (!length(method_forms)) {
    method_forms <- structure(list(), names = character())
  }

  # expectation that the names are the same (possibly redundant but quicker to)
  bad <- waldo::compare(
    x     = method_args,
    y     = setdiff(fun_args, "wb"),
    x_arg = method0,
    y_arg = fun
  )

  if (length(bad)) {
    testthat::fail(bad)
    return(invisible())
  }

  # expectation that the default values are the same
  bad <- waldo::compare(
    x     = method_forms,
    y     = fun_forms,
    x_arg = method0,
    y_arg = fun
  )

  if (length(bad)) {
    testthat::fail(bad)
    return(invisible())
  }

  if (!is.null(params)) {
    # create now so that it's the same every time
    wb_fun <- wb$clone(deep = TRUE)
    wb_method <- wb$clone(deep = TRUE)

    # be careful and report when we failed to run these


    # the style names are generated at random: use matching seeds for both calls
    options("openxlsx2_seed" = NULL)
    if (ignore_wb) {
      res_fun    <- try(do.call(fun, c(params)), silent = TRUE)
    } else {
      res_fun    <- try(do.call(fun, c(wb = wb_fun, params)), silent = TRUE)
    }

    options("openxlsx2_seed" = NULL)
    res_method <- try(do.call(wb_method[[method]], params), silent = TRUE)

    msg <- NULL

    if (inherits(res_fun, "try-error")) {
      temp <- file()
      writeLines(attr(res_fun, "condition")$message, temp)
      bad <- paste0("# > ", readLines(temp))
      close(temp)
      expr <- paste0("\n", deparse1(do.call(call, c(fun, c(wb = wb_fun, params)))))
      msg <- c(msg, expr, bad)
    }

    if (inherits(res_method, "try-error")) {
      temp <- file()
      writeLines(attr(res_method, "condition")$message, temp)
      bad <- paste0("# > ", readLines(temp))
      close(temp)
      expr <- paste0("\n", deparse1(do.call(call, c(method0, params))))
      # slight fix because we get those "`" which I don't want to see
      expr <- sub(paste0("`", method0, "`"), method0, expr, fixed = TRUE)
      msg <- c(msg, expr, bad)
    }

    if (!is.null(msg)) {
      # browser()
      bad <- c("Failed to get results", msg)
      testthat::fail(bad)
      return(invisible())
    }

    for (i in ignore_fields) {
      # remove the fields we don't want to check
      res_method[[i]] <- NULL
      res_fun[[i]]    <- NULL
    }

    # expectation that the results are the same
    bad <- waldo::compare(
      x                  = res_method,
      y                  = res_fun,
      x_arg              = method0,
      y_arg              = fun,
      ignore_attr        = ignore_attr,
      ignore_formula_env = TRUE
    )

    if (length(bad)) {
      testthat::fail(bad)
      return(invisible())
    }
  }

  testthat::succeed()
  return(invisible())
}

#' @description
#' A trimmed down pseudo wrapper to check internal wrapped functions.
#' @rdname expect_wrapper
#' @keywords internal
#' @noRd
expect_pseudo_wrapper <- function(
    method,
    fun           = paste0("wb_", method)
) {
  method_fun <- get(method, wbWorkbook$public_methods)
  fun_fun    <- match.fun(fun)

  method_forms <- as.list(formals(method_fun))
  fun_forms    <- as.list(formals(fun_fun))

  method_args <- names(method_forms)
  fun_args    <- names(fun_forms)

  ignore <- "xlsxFile"

  # remove ignores from fun
  m <- match(ignore, fun_args, 0L)
  if (!identical(m, 0L)) {
    # remove wb from both the args and formals
    fun_args  <- fun_args[-m]
    fun_forms <- fun_forms[-m]
  }

  # remove ignores from method
  m <- match(ignore, method_args, 0L)
  if (!identical(m, 0L)) {
    # remove wb from both the args and formals
    method_args  <- method_args[-m]
    method_forms <- method_forms[-m]
  }

  # expectation that the names are the same (possibly redundant but quicker to)
  bad <- waldo::compare(
    x     = method_args,
    y     = setdiff(fun_args, "wb")
  )

  if (length(bad)) {
    testthat::fail(bad)
    return(invisible())
  }

  # expectation that the default values are the same
  bad <- waldo::compare(
    x     = method_forms,
    y     = fun_forms
  )

  if (length(bad)) {
    testthat::fail(bad)
    return(invisible())
  }

  testthat::succeed()
  return(invisible())
}
