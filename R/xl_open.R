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


#' Open a file or workbook object in spreadsheet software
#'
#' @description
#' `xl_open()` is a portable utility designed to open spreadsheet files (such
#' as .xlsx or .xls) or [wbWorkbook] objects using the appropriate application
#' based on the host operating system. It handles the nuances of background
#' execution to ensure the R interpreter remains unblocked.
#'
#' @details
#' The method for identifying and launching the software varies by platform:
#'
#' **Windows** utilizes `shell.exec()` to trigger the file association
#' registered with the operating system.
#'
#' **macOS** utilizes the system `open` command, which respects default
#' application handlers for the file type. Users can override the default by
#' setting `options("openxlsx2.excelApp")`.
#'
#' **Linux** attempts to locate common spreadsheet utilities in the system
#' path, including LibreOffice (`soffice`), Gnumeric (`gnumeric`), Calligra
#' Sheets (`calligrasheets`), and ONLYOFFICE (`onlyoffice-desktopeditors`).
#' If multiple applications are found during an interactive session, a menu is
#' presented to the user to define their preference, which is then stored in
#' `options("openxlsx2.excelApp")`.
#'
#' For `wbWorkbook` objects, the function automatically clones the workbook,
#' detects the presence of macros (VBA) to determine the appropriate temporary
#' file extension, and saves the content to a temporary location before opening.
#'
#' @param x A character string specifying the path to a spreadsheet file or
#'   a [wbWorkbook] object.
#' @param interactive Logical; if `FALSE`, the function will not attempt to
#'   launch the application. Defaults to the result of [base::interactive()].
#' @param flush Logical; if `TRUE`, the workbook is written to the temporary
#'   location using the stream-based XML parser. See [wb_save()] for details.
#'
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
