#' Open a Microsoft Excel file (xls/xlsx) or an openxlsx2 wbWorkbook
#'
#' @description This function tries to open a Microsoft Excel (xls/xlsx) file or
#'   an openxlsx2 wbWorkbook with the proper application, in a portable manner.
#'
#'   In Windows it uses `base::shell.exec()` (Windows only function) to
#'   determine the appropriate program.
#'
#'   In Mac (c) it uses system default handlers, given the file type.
#'
#'   In Linux it searches (via `which`) for available xls/xlsx reader
#'   applications (unless `options('openxlsx2.excelApp')` is set to the app bin
#'   path), and if it finds anything, sets `options('openxlsx2.excelApp')` to the
#'   program chosen by the user via a menu (if many are present, otherwise it
#'   will set the only available). Currently searched for apps are
#'   Libreoffice/Openoffice (`soffice` bin), Gnumeric (`gnumeric`) and Calligra
#'   Sheets (`calligrasheets`).
#'
#' @param x A path to the Excel (xls/xlsx) file or Workbook object.
#' @param interactive If `FALSE` will throw a warning and not open the path.
#'   This can be manually set to `TRUE`, otherwise when `NA` (default) uses the
#'   value returned from [base::interactive()]
#' @export
#' @examples
#' \donttest{
#' if (interactive()) {
#'   xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#'   xl_open(xlsxFile)
#'
#'   # (not yet saved) Workbook example
#'   wb <- wb_workbook()
#'   x <- mtcars[1:6, ]
#'   wb$add_worksheet("Cars")
#'   wb$add_data("Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#'   xl_open(wb)
#' }
#' }
xl_open <- function(x, interactive = NA) {
  UseMethod("xl_open")
}

#' @rdname xl_open
#' @export
xl_open.wbWorkbook <- function(x, interactive = NA) {
  stopifnot(R6::is.R6(x))
  has_macros <- isTRUE(length(x$vbaProject) > 0)
  xl_open(x$clone()$save(temp_xlsx(macros = has_macros))$path, interactive = interactive)
}

#' @rdname xl_open
#' @export
xl_open.default <- function(x, interactive = NA) {
  stopifnot(file.exists(x))

  # nocov start
  if (is.na(interactive)) {
    interactive <- interactive()
  }
  # nocov end

  if (!isTRUE(interactive)) {
    warning("will not open file when not interactive")
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
      system2('open', shQuote(file))
    },
    stop("Operating system not handled: ", toString(userSystem))
  )
  # nocov end
}


chooseExcelApp <- function() {

  # nocov start
  m <- c(
    `Libreoffice/OpenOffice` = "soffice",
    `Calligra Sheets` = "calligrasheets",
    `Gnumeric` = "gnumeric"
  )

  prog <- Sys.which(m)
  names(prog) <- names(m)
  availProg <- prog["" != prog]
  nApps <- length(availProg)

  if (0 == nApps) {
    stop(
      "No applications (detected) available.\n",
      "Set options('openxlsx.excelApp'), instead."
    )
  }

  if (1 == nApps) {
    message("Only ", names(availProg), " found")
    unnprog <- unname(availProg)
    message(sprintf("Setting options(openxlsx2.excelApp = '%s')", unnprog))
    options(openxlsx2.excelApp = unnprog)
    return(invisible(unnprog))
  }

  if (1 < nApps) {
    if (!interactive()) {
      stop(
        "Cannot choose an Excel file opener non-interactively.\n",
        "Set options('openxlsx.excelApp'), instead."
      )
    }
    res <- menu(names(availProg), title = "Excel Apps availables")
    unnprog <- unname(availProg[res])
    if (res > 0L) {
      message(sprintf("Setting options(openxlsx2.excelApp = '%s')", unnprog))
      options(openxlsx2.excelApp = unnprog)
    }
    return(invisible(unname(unnprog)))
  }
  # nocov end

  stop("Unexpected error in openxlsx2::chooseExcelApp()") # nocov

}
