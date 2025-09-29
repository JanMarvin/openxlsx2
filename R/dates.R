# `convert_date()` ------------------------------------------------
#' Convert from Excel date, datetime or hms number to R Date type
#'
#' Convert from Excel date number to R Date type
#'
#' Excel stores dates as number of days from some origin day
#'
#' @seealso [wb_add_data()]
#' @param x A vector of integers
#' @param origin date. Default value is for Windows Excel 2010
#' @param tz A timezone, defaults to "UTC"
#' @return A date, datetime, or hms.
#' @details Setting the timezone in [convert_datetime()] will alter the value. If users expect a datetime value in a specific timezone, they should try e.g. `lubridate::force_tz`.
#' @export
#' @examples
#' # date --
#' ## 2014 April 21st to 25th
#' convert_date(c(41750, 41751, 41752, 41753, 41754, NA))
#' convert_date(c(41750.2, 41751.99, NA, 41753))
#'
#' # datetime --
#' ##  2014-07-01, 2014-06-30, 2014-06-29
#' x <- c(41821.8127314815, 41820.8127314815, NA, 41819, NaN)
#' convert_datetime(x)
#' convert_datetime(x, tz = "Australia/Perth")
#' convert_datetime(x, tz = "UTC")
#'
#' # hms ---
#' ## 12:13:14
#' x <- 0.50918982
#' convert_hms(x)
convert_date <- function(x, origin = "1900-01-01") {
  sel <- is_charnum(x)
  out <- x
  out[sel] <- date_to_unix(x[sel], origin = origin, datetime = FALSE)
  out <- as.double(out)
  .Date(out)
}

#' @rdname convert_date
#' @export
convert_datetime <- function(x, origin = "1900-01-01", tz = "UTC") {
  sel <- is_charnum(x)
  out <- x
  out[sel] <- date_to_unix(x[sel], origin = origin, datetime = TRUE)
  out <- as.double(out)
  .POSIXct(out, tz)
}

#' @rdname convert_date
#' @export
convert_hms <- function(x) {
  if (isNamespaceLoaded("hms")) {
    x <- convert_datetime(x, origin = "1970-01-01", tz = "UTC")
    class(x) <- c("hms", "difftime")
  } else {
    x <- convert_datetime(x, origin = "1970-01-01")
    x <- format(x, format = "%H:%M:%S")
  }
  x
}

# `convert_date()` helpers --------------------
#' Get the date origin an xlsx file is using
#'
#' Return the date origin used internally by an xlsx or xlsm file
#'
#' Excel stores dates as the number of days from either 1904-01-01 or 1900-01-01. This function
#' checks the date origin being used in an Excel file and returns is so it can be used in [convert_date()]
#'
#' @param xlsxFile An xlsx or xlsm file or a wbWorkbook object.
#' @param origin return the origin instead of the character string.
#' @return One of "1900-01-01" or "1904-01-01".
#' @seealso [convert_date()]
#' @examples
#' ## create a file with some dates
#' temp <- temp_xlsx()
#' write_xlsx(as.Date("2015-01-10") - (0:4), file = temp)
#' m <- read_xlsx(temp)
#'
#' ## convert to dates
#' do <- get_date_origin(system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))
#' convert_date(m[[1]], do)
#'
#' get_date_origin(wb_workbook())
#' get_date_origin(wb_workbook(), origin = TRUE)
#' @noRd
get_date_origin <- function(xlsxFile, origin = FALSE) {

  # TODO: allow using a workbook?
  if (!inherits(xlsxFile, "wbWorkbook"))
    xlsxFile <- wb_load(xlsxFile)

  if (grepl('date1904="1"|date1904="true"', xlsxFile$workbook$workbookPr, ignore.case = TRUE)) {
    z <- ifelse(origin, 24107L, "1904-01-01")
  } else {
    z <- ifelse(origin, 25569L, "1900-01-01")
  }

  z
}

parseOffset <- function(tz) {
  # likely for cases where tz is NA
  z <- vector("numeric", length = length(tz))
  z[is.na(tz)] <- NA_real_
  tz <- tz[!is.na(tz)]

  z[!is.na(z)] <- ifelse(stringi::stri_sub(tz, 1, 1) == "+", 1L, -1L) *
    (as.integer(stringi::stri_sub(tz, 2, 3)) +
     as.integer(stringi::stri_sub(tz, 4, 5)) / 60) / 24

  ifelse(is.na(z), 0, z)
}

#' @name as_POSIXct_utc
#' @title Convert to POSIXct with timezone UTC
#' @param x something as.POSIXct can convert
#' @noRd
as_POSIXct_utc <- function(x) {
  z <- as.POSIXct(x, tz = "UTC")
  # this should be superfluous?
  attr(z, "tzone") <- "UTC"
  z
}

#' helper function to convert hms to posix
#' @param x a hms object
#' @noRd
as_POSIXlt_hms <- function(x) {
  z <- as.POSIXlt("1970-01-01")
  units(x) <- "secs"
  z$sec <- as.numeric(x)
  z
}

# `convert_to_excel_date()` helpers -----------------------------------
#' conversion helper function
#' @param x a date, posixct or hms object
#' @param date1904 a special time format in openxml
#' @noRd
conv_to_excel_date <- function(x, date1904 = FALSE) {

  is_hms <- inherits(x, "hms")
  to_convert  <- inherits(x, "POSIXlt") || inherits(x, "Date") || is_hms
  if (to_convert) {
    # as.POSIXlt does not use local timezone
    if (inherits(x, "Date")) x <- as.POSIXlt(x)
    if (is_hms) {
      class(x) <- c("hms", "difftime")
      # helper function if other conversion function is available.
      # e.g. when testing without hms loaded
      x <- tryCatch({
        as.POSIXlt(x)
      }, error = function(e) {
          as_POSIXlt_hms(x)
      })
    }
    x <- as.POSIXct(x, tz = "UTC")
  }

  if (!inherits(x, "Date") && !inherits(x, "POSIXct")) {
    warning("could not convert x to Excel date. x is of class: ",
            class(x))
    return(x)
  }

  ## convert any Dates to integers and create date style object
  origin <- 25569L
  if (date1904) origin <- 24107L
  if (is_hms) origin <- 0

  if (inherits(x, "POSIXct")) {
    tz <- format(x, "%z")
    offSet <- parseOffset(tz)

    z <- as.numeric(x) / 86400L + origin + offSet

    # Excel ships a date 29Feb1900 due to initial backward comparability
    # with previous spreadsheet software.
    if (origin == 25569L) {
      earlyDate <- z <= 60L & !is.na(z)
      z[earlyDate] <- z[earlyDate] - 1L
    }
  }

  if (!is_hms && any(z < 1, na.rm = TRUE)) {
    warning("Date < 1900-01-01 found. This can not be converted.")
  }

  z
}

# `convert_to_excel_date()` ---------------------------
#' convert back to an Excel Date
#' @param df dataframe
#' @param date1904 take different origin
#' @examples
#'  xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#'  wb1 <- wb_load(xlsxFile)
#'  df <- wb_to_df(wb1)
#'  # conversion is done on dataframes only
#'  convert_to_excel_date(df = df["Var5"], date1904 = FALSE)
#' @export
convert_to_excel_date <- function(df, date1904 = FALSE) {

  is_date <- vapply(
    df,
    function(x) {
      inherits(x, "Date") || inherits(x, "POSIXct") || inherits(x, "hms")
    },
    NA
  )

  df[is_date] <- lapply(df[is_date], FUN = conv_to_excel_date, date1904 = date1904)

  df
}
