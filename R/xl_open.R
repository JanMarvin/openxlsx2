chooseExcelApp <- function() {

  # nocov start
  m <- c(
    `Libreoffice/OpenOffice` = "soffice",
    `Calligra Sheets` = "calligrasheets",
    `Gnumeric` = "gnumeric",
    `ONLYOFFICE` = "onlyoffice-desktopeditors"
  )

  prog <- Sys.which(m)
  names(prog) <- names(m)
  availProg <- prog["" != prog]
  nApps <- length(availProg)

  if (0 == nApps) {
    stop(
      "No applications (detected) available.\n",
      "Set options('openxlsx2.excelApp'), instead."
    )
  }

  ## TODO previously openxlsx/openxlsx2 were messaging this. Still needed?
  # message("Only ", names(availProg), " found")
  # unnprog <- unname(availProg)
  # message(sprintf("Setting options(openxlsx2.excelApp = '%s')", unnprog))
  if (1 == nApps) {
    unnprog <- unname(availProg)
    options(openxlsx2.excelApp = unnprog)
    return(invisible(unnprog))
  }

  if (1 < nApps) {
    if (!interactive()) {
      stop(
        "Cannot choose a spreadsheet file opener non-interactively.\n",
        "Set options('openxlsx2.excelApp'), instead."
      )
    }
    res <- menu(names(availProg), title = "Spreadsheet software available")
    unnprog <- unname(availProg[res])
    if (res > 0L) {
      message(sprintf("Setting options(openxlsx2.excelApp = '%s')", unnprog))
      options(openxlsx2.excelApp = unnprog)
    }
    return(invisible(unname(unnprog)))
  }
  # nocov end

  stop("Unexpected error in openxlsx2:::chooseExcelApp()") # nocov

}


#' Open an xlsx file or a `wbWorkbook` object
#'
#' @description
#' This function tries to open a Microsoft Excel (xls/xlsx) file or,
#' an [openxlsx2::wbWorkbook] with the proper application, in a portable manner.
#'
#' On Windows it uses `base::shell.exec()` (Windows only function) to
#' determine the appropriate program.
#'
#' On Mac, (c) it uses system default handlers, given the file type.
#'
#' On Linux, it searches (via `which`) for available xls/xlsx reader
#' applications (unless `options('openxlsx2.excelApp')` is set to the app bin
#' path), and if it finds anything, sets `options('openxlsx2.excelApp')` to the
#' program chosen by the user via a menu (if many are present, otherwise it
#' will set the only available). Currently searched for apps are
#' Libreoffice/Openoffice (`soffice` bin), Gnumeric (`gnumeric`), Calligra
#' Sheets (`calligrasheets`) and ONLYOFFICE (`onlyoffice-desktopeditors`).
#'
#' @param x A path to a spreadsheet file or wbWorkbook object. This can be any
#'   file type that can be opened in the corresponding software.
#' @param interactive If `FALSE` will throw a warning and not open the path.
#'   This can be manually set to `TRUE`, otherwise when `NA` (default) uses the
#'   value returned from [base::interactive()]
#' @param flush If `TRUE` the `flush` argument of [wb_save()] will be used to
#' create the output file. Applies only to workbooks.
#' @examples
#' \donttest{
#' if (interactive()) {
#'   xlsx_file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#'   xl_open(xlsx_file)
#'
#'   # (not yet saved) Workbook example
#'   wb <- wb_workbook()
#'   x <- mtcars[1:6, ]
#'   wb$add_worksheet("Cars")
#'   wb$add_data("Cars", x, start_col = 2, start_row = 3, row_names = TRUE)
#'   xl_open(wb)
#' }
#' }
#' @export
xl_open <- function(x, interactive = NA, flush = FALSE) {
  # The only function to accept a workbook or a file.
  UseMethod("xl_open")
}

#' @rdname xl_open
#' @export
xl_open.wbWorkbook <- function(x, interactive = NA, flush = FALSE) {
  assert_workbook(x)
  has_macros <- isTRUE(length(x$vbaProject) > 0)
  xl_open(x$clone()$save(temp_xlsx(macros = has_macros), flush = flush)$path, interactive = interactive)
}

#' @rdname xl_open
#' @export
xl_open.default <- function(x, interactive = NA, flush = FALSE) {
  stopifnot(file.exists(x))

  # nocov start
  if (is.na(interactive)) {
    interactive <- interactive()
  }
  # nocov end

  if (!isTRUE(interactive)) {
    warning("will not open file when not interactive", call. = FALSE)
    return()
  }

  # nocov start

  ## execution should be in background in order to not block R
  ## interpreter
  file <- normalizePath(x, mustWork = TRUE)
  userSystem <- Sys.info()["sysname"]

  switch(
    userSystem,
    Linux = {
      app <- getOption("openxlsx2.excelApp", chooseExcelApp())
      system2(app, c(file, "&"))
    },
    Windows = {
      shell.exec(file) # nolint
    },
    Darwin = {
      # system2('/Applications/LibreOffice.app/Contents/MacOS/soffice', shQuote(file), wait = FALSE)
      if (!is.null(getOption("openxlsx2.excelApp")))
        system2('open', paste0("-a \"", getOption("openxlsx2.excelApp"), "\" ", shQuote(file)))
      else
        system2('open', shQuote(file))
    },
    stop("Operating system not handled: ", toString(userSystem))
  )
  # nocov end
}
