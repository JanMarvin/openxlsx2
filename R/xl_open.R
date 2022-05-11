#' Open a Microsoft Excel file (xls/xlsx) or an openxlsx Workbook
#'
#' @description This function tries to open a Microsoft Excel (xls/xlsx) file or
#'   an openxlsx Workbook with the proper application, in a portable manner.
#'
#'   In Windows it uses [base::shell.exec()] to determine the appropriate
#'   program.
#'
#'   In Mac (c) it uses system default handlers, given the file type.
#'
#'   In Linux it searches (via `which`) for available xls/xlsx reader
#'   applications (unless `options('openxlsx.excelApp')` is set to the app bin
#'   path), and if it finds anything, sets `options('openxlsx.excelApp')` to the
#'   program chosen by the user via a menu (if many are present, otherwise it
#'   will set the only available). Currently searched for apps are
#'   Libreoffice/Openoffice (`soffice` bin), Gnumeric (`gnumeric`) and Calligra
#'   Sheets (`calligrasheets`).
#'
#' @param file path to the Excel (xls/xlsx) file or Workbook object.
#' @param interactive If `FALSE` will throw a warning and not open the path.
#'   This can be manually set to `TRUE`, otherwise when `NA` (defualt) uses the
#'   value returned from [base::interactive()]
#' @export
#' @examples
#' # file example
#' example(write_data)
#' # xl_open("write_dataExample.xlsx")
#'
#' # (not yet saved) Workbook example
#' wb <- wb_workbook()
#' x <- mtcars[1:6, ]
#' wb$add_worksheet("Cars")
#' wb$add_data("Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#' xl_open(wb)
xl_open <- function(file = NULL, interactive = NA) {
  UseMethod("xl_open")
}

#' @rdname xl_open
#' @export
xl_open.wbWorkbook <- function(file = NULL, interactive = NA) {
  stopifnot(R6::is.R6(file))
  xl_open(file$clone()$save(temp_xlsx())$path)
}

#' @rdname xl_open
#' @export
xl_open.default <- function(file = NULL, interactive = NA) {
  if (is.na(interactive)) {
    interactive <- interactive()
  }

  if (!isTRUE(interactive)) {
    warning("will not open file when not interactive")
    return()
  }

  stopifnot(!is.null(file), file.exists(file))

  ## execution should be in background in order to not block R
  ## interpreter
  file <- normalizePath(file, mustWork = TRUE)
  userSystem <- Sys.info()["sysname"]

  switch(
    userSystem,
    Linux = {
      app <- getOption("openxlsx2.excelApp", chooseExcelApp())
      myCommand <- paste(app, file, "&", sep = " ")
      system(command = myCommand)
    },
    Windows = {
      shell.exec(file)
    },
    Darwin = {
      myCommand <- paste0('open ', shQuote(file))
      system(command = myCommand)
    },
    stop("Operating system not handled: ", toString(userSystem))
  )
}


chooseExcelApp <- function() {
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
    invisible(unnprog)
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
    invisible(unname(unnprog))
  }

  stop("Unexpected error in openxlsx2::chooseExcelApp()")
}
