#' @name as_POSIXct_utc
#' @title Convert to POSIXct with timezone UTC
#' @param x something as.POSIXct can convert
#' @keywords internal
#' @noRd
as_POSIXct_utc <- function(x) {
  z <- as.POSIXct(x, tz = "UTC")
  attr(z, "tzone") <- "UTC"
  z
}


#' @name convert_date
#' @title Convert from excel date number to R Date type
#' @description Convert from excel date number to R Date type
#' @param x A vector of integers
#' @param origin date. Default value is for Windows Excel 2010
#' @param ... additional parameters passed to as.Date()
#' @details Excel stores dates as number of days from some origin day
#' @seealso [write_data()]
#' @export
#' @examples
#' ## 2014 April 21st to 25th
#' convert_date(c(41750, 41751, 41752, 41753, 41754, NA))
#' convert_date(c(41750.2, 41751.99, NA, 41753))
convert_date <- function(x, origin = "1900-01-01", ...) {

  # use as.integer to only get the integer part of a number.
  # in openxml dates are integers only.
  x <- as.integer(x)
  notNa <- !is.na(x)
  earlyDate <- x < 60

  if (origin == "1900-01-01") {
    x[notNa] <- x[notNa] - 2
    x[earlyDate & notNa] <- x[earlyDate & notNa] + 1
  }

  as.Date(x, origin = origin, ...)
}


#' @name convert_datetime
#' @title Convert from excel time number to R POSIXct type.
#' @description Convert from excel time number to R POSIXct type.
#' @param x A numeric vector
#' @param origin date. Default value is for Windows Excel 2010
#' @param ... Additional parameters passed to as.POSIXct
#' @details Excel stores dates as number of days from some origin date
#' @export
#' @examples
#' ## 2014-07-01, 2014-06-30, 2014-06-29
#' x <- c(41821.8127314815, 41820.8127314815, NA, 41819, NaN)
#' convert_datetime(x)
#' convert_datetime(x, tz = "Australia/Perth")
#' convert_datetime(x, tz = "UTC")
convert_datetime <- function(x, origin = "1900-01-01", ...) {
  x <- as.numeric(x)
  date <- convert_date(x, origin)

  x <- x * 86400
  rem <- x %% 86400

  hours <- as.integer(floor(rem / 3600))
  minutes_fraction <- rem %% 3600
  minutes_whole <- as.integer(floor(minutes_fraction / 60))
  secs <- minutes_fraction %% 60

  y <- sprintf("%02d:%02d:%06.3f", hours, minutes_whole, secs)
  notNA <- !is.na(x)
  date_time <- rep(NA, length(x))
  date_time[notNA] <- as.POSIXct(paste(date[notNA], y[notNA]), ...)

  .POSIXct(date_time)
}



#' @name get_date_origin
#' @title Get the date origin an xlsx file is using
#' @description Return the date origin used internally by an xlsx or xlsm file
#' @param xlsxFile An xlsx or xlsm file or a wbWorkbook object.
#' @param origin return the origin instead of the character string.
#' @details Excel stores dates as the number of days from either 1904-01-01 or 1900-01-01. This function
#' checks the date origin being used in an Excel file and returns is so it can be used in [convert_date()]
#' @return One of "1900-01-01" or "1904-01-01".
#' @seealso [convert_date()]
#' @examples
#'
#' ## create a file with some dates
#' temp <- temp_xlsx()
#' write_xlsx(as.Date("2015-01-10") - (0:4), file = temp)
#' m <- read_xlsx(temp)
#'
#' ## convert to dates
#' do <- get_date_origin(system.file("extdata", "readTest.xlsx", package = "openxlsx2"))
#' convert_date(m[[1]], do)
#'
#' get_date_origin(wb_workbook())
#' get_date_origin(wb_workbook(), origin = TRUE)
#' @export
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

conv_to_excel_date <- function(x, date1904 = FALSE) {

  to_convert <- inherits(x, "POSIXlt") || inherits(x, "Date")
  if (to_convert) {
    # as.POSIXlt does not use local timezone
    if (inherits(x, "Date")) x <- as.POSIXlt(x)
    x <- as.POSIXct(x)
  }

  if (!inherits(x, "Date") && !inherits(x, "POSIXct")) {
    warning("could not convert x to Excel date. x is of class: ",
            class(x))
    return(x)
  }

  ## convert any Dates to integers and create date style object
  origin <- 25569L
  if (date1904) origin <- 24107L

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

  if (any(z < 1, na.rm = TRUE)) {
    warning("Date < 1900-01-01 found. This can not be converted.")
  }

  z
}

#' convert back to ExcelDate
#' @param df dataframe
#' @param date1904 take different origin
#' @examples
#'  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
#'  wb1 <- wb_load(xlsxFile)
#'  df <- wb_to_df(wb1)
#'  # conversion is done on dataframes only
#'  convertToExcelDate(df = df["Var5"], date1904 = FALSE)
#' @export
convertToExcelDate <- function(df, date1904 = FALSE) {

  df_class <- as.data.frame(Map(class, df))
  is_date <- apply(df_class, 2, function(x) any(x %in%  c("Date", "POSIXct")))

  df[is_date] <- lapply(df[is_date], FUN = conv_to_excel_date, date1904 = date1904)

  df
}
