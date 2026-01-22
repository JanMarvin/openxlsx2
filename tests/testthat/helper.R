expect_equal_workbooks <- function(object, expected, ..., ignore_fields = NULL) {
  assert_workbook(object)
  assert_workbook(expected)

  # Check for invalid fields
  fields <- c(names(wbWorkbook$public_fields), names(wbWorkbook$private_fields))
  bad_fields <- setdiff(ignore_fields, fields)
  if (length(bad_fields)) stop("Invalid fields: ", toString(bad_fields))

  # Nullify ignored fields
  for (i in ignore_fields) {
    object[[i]] <- NULL
    expected[[i]] <- NULL
  }

  diffs <- all.equal(object, expected, ...)

  if (!isTRUE(diffs)) {
    testthat::fail(paste(diffs, collapse = "\n"))
  } else {
    testthat::succeed()
  }
  invisible()
}

#' Expectation wrapper for wbWorkbook methods
#'
#' Internal tool used to ensure that `wbWorkbook` methods and their functional
#' wrappers (e.g., `wb_method()`) remain synchronized in terms of arguments,
#' default values, and execution output.
#'
#' @param method The name of the [wbWorkbook] public method.
#' @param fun The name of the wrapper function. Defaults to `wb_method`.
#' @param wb A `wbWorkbook` object to use for testing.
#' @param params A named list of parameters to pass to both functions.
#'   If `NULL`, only the function signatures (formals) are compared.
#' @param ignore Names of parameters to ignore during the signature check.
#' @param ignore_fields Internal workbook fields to remove from the objects
#'   before comparing results.
#' @param ignore_wb Boolean. Set to `TRUE` for pseudo-wrappers (like `wb_load`)
#'   that do not take a workbook as the first argument.
#' @returns Invisibility, called for its side effects in `testthat`.
expect_wrapper <- function(
    method,
    fun           = paste0("wb_", method),
    wb            = wb_workbook(),
    params        = NULL,
    ignore        = NULL,
    ignore_fields = NULL,
    ignore_wb     = FALSE
) {

  assert_workbook(wb)
  if (!is.character(method) || length(method) != 1L) {
    stop("`method` must be a single string representing the R6 method name.")
  }
  if (!is.character(fun) || length(fun) != 1L) {
    stop("`fun` must be a single string representing the wrapper function name.")
  }
  if (!is.null(params) && !is.list(params)) {
    testthat::fail(sprintf("Test config error: `params` for '%s' must be a list or NULL.", method))
    return(invisible())
  }

  method_fun <- get(method, wbWorkbook$public_methods)
  fun_fun    <- match.fun(fun)

  m_forms <- as.list(formals(method_fun))
  f_forms <- as.list(formals(fun_fun))

  to_ignore <- unique(c("wb", "file", ignore))
  m_forms   <- m_forms[!names(m_forms) %in% to_ignore]
  f_forms   <- f_forms[!names(f_forms) %in% to_ignore]

  # 1. Compare Signatures (Formals)
  diffs <- all.equal(m_forms, f_forms, check.attributes = FALSE)

  if (!isTRUE(diffs)) {
    testthat::fail(paste("Formals mismatch for", method, ":", paste(diffs, collapse = "\n")))
    return(invisible())
  }

  # 2. Compare Execution Results
  if (!is.null(params)) {
    wb_fun     <- wb$clone(deep = TRUE)
    wb_method  <- wb$clone(deep = TRUE)
    wb_fun_pre <- wb_fun$clone(deep = TRUE)

    options("openxlsx2_seed" = 123)
    res_fun <- if (ignore_wb) {
      try(do.call(fun, params), silent = TRUE)
    } else {
      try(do.call(fun, c(list(wb = wb_fun), params)), silent = TRUE)
    }

    options("openxlsx2_seed" = 123)
    res_method <- try(do.call(wb_method[[method]], params), silent = TRUE)

    if (inherits(res_fun, "try-error") || inherits(res_method, "try-error")) {
      testthat::fail(paste("Execution failed for", method))
      return(invisible())
    }

    for (i in ignore_fields) {
      res_method[[i]] <- res_fun[[i]] <- NULL
    }

    res_diffs <- all.equal(res_method, res_fun, check.names = TRUE)
    if (!isTRUE(res_diffs)) {
      testthat::fail(paste("Result mismatch:", paste(res_diffs, collapse = "\n")))
    }

    mut_diffs <- all.equal(wb_fun, wb_fun_pre, check.attributes = FALSE)
    if (!isTRUE(mut_diffs)) {
      testthat::fail(paste(fun, "mutated the input workbook"))
    }
  }

  testthat::succeed()
  invisible()
}

# Miscellaneous helpers for testthat -----------

#' provides testfile path for testthat
#' @param x a file assumed in testfiles folder
#' @param replace logical; if TRUE, re-downloads the file even if it exists
testfile_path <- function(x, replace = FALSE) {

  test_dir <- testthat::test_path("testfiles")
  if (!dir.exists(test_dir)) dir.create(test_dir, recursive = TRUE)

  fl <- file.path(test_dir, x)

  if (Sys.getenv("openxlsx2_testthat_fullrun") == "") {
    if (isTRUE(as.logical(Sys.getenv("CI", "false")))) # on_ci()
      return(testthat::skip("Skip on CI"))

    if (!interactive() && !isTRUE(as.logical(Sys.getenv("NOT_CRAN", "false")))) # on_cran()
      return(testthat::skip("Skip on CRAN"))
  }

  # Download logic
  if (!file.exists(fl) || replace) {
    url <- paste0("https://github.com/JanMarvin/openxlsx-data/raw/main/", x)
    tryCatch({
      download.file(url, destfile = fl, quiet = TRUE, mode = "wb")
    }, error = function(e) NULL)
  }

  if (!file.exists(fl)) {
    testthat::skip(paste("Testfile", x, "could not be acquired."))
  }

  fl
}

testsetup <- function() {
  if (is.null(getOption("openxlsx2.datetimeCreated"))) {
    testthat::test_that("testsetup", {
      options("openxlsx2.datetimeCreated" = as.POSIXct("2023-07-20 23:32:14", tz = "UTC"))
      exp <- "2023-07-20T23:32:14Z"
      got <- wb_workbook()$get_properties()[["datetime_created"]]
      testthat::expect_equal(got, exp)
    })
  }
}

#' @param host character; host to ping
dns_lookup <- function(host = "captive.apple.com") {
  con <- try(socketConnection(host, port = 80, open = "r+", timeout = 2), silent = TRUE)
  if (inherits(con, "connection")) {
    on.exit(close(con))
    return(TRUE)
  }
  FALSE
}

#' Skip tests if offline or on CRAN
skip_online_checks <- function() {
  testthat::skip_on_cran()
  if (!dns_lookup()) {
    testthat::skip("Offline: DNS lookup failed")
  }
}
