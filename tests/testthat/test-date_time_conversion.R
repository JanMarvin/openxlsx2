
test_that("as_POSIXct_utc", {
  exp <- "2022-03-02 19:27:35"
  got <- as_POSIXct_utc("2022-03-02 19:27:35")
  expect_equal(exp, as.character(got))
})

test_that("convert to date", {
  dates <- as.Date("2015-02-07") + -10:10
  origin <- 25569L
  n <- as.integer(dates) + origin

  expect_equal(convert_date(n), dates)

  earlyDate <- as.Date("1900-01-03")
  serialDate <- 3
  expect_equal(convert_date(serialDate), earlyDate)

})

test_that("convert to datetime", {
  x <- 43037 + 2 / 1440
  res <- convert_datetime(x, tz = Sys.timezone())
  exp <- as.POSIXct(1509235320, tz = Sys.timezone(), origin = "1970-01-01")
  expect_equal(res, exp, ignore_attr = "tzone")

  x <- 43037 + 2 / 1440 + 1 / 86400
  res <- convert_datetime(x, tz = Sys.timezone())
  exp <- as.POSIXct(1509235321, tz = Sys.timezone(), origin = "1970-01-01")
  expect_equal(res, exp, ignore_attr = "tzone")

  x <- 43037 + 2.50 / 1440
  res <- convert_datetime(x, tz = Sys.timezone())
  exp <- as.POSIXct(1509235350, tz = Sys.timezone(), origin = "1970-01-01")
  expect_equal(res, exp, ignore_attr = "tzone")

  x <- 43037 + 2 / 1440 + 12.12 / 86400
  x_datetime <- convert_datetime(x, tz = "UTC")
  attr(x_datetime, "tzone") <- "UTC"
  exp <- as.POSIXct("2017-10-29 00:02:12", tz = "UTC", origin = "1970-01-01")
  expect_equal(res, exp, ignore_attr = "tzone")
})

test_that("convert hms works", {

  x <- structure(43994, units = "secs", class = c("hms", "difftime"))

  wb <- wb_workbook()$add_worksheet()$add_data(x = x, col_names = FALSE)

  exp <- data.frame(
    r = "A1", row_r = "1", c_r = "A", c_s = "1", c_t = "",
    v = "0.509189814814815",
    f = "", f_attr = "", is = "",
    stringsAsFactors = FALSE)
  got <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(exp, got)

  z <- wb_to_df(wb, col_names = FALSE, keep_attributes = TRUE)
  expect_equal(z$A, "12:13:14")
  expect_equal(attr(z, "tt")$A, "h")

})

test_that("custom classes are treated independently", {

  skip_on_cran()

  # create a custom test class
  as.character.myclass <- function(x, ...) paste("myclass:", format(x, digits = 2))
  assign("as.character.myclass", as.character.myclass, envir = globalenv())
  on.exit(rm("as.character.myclass", envir = globalenv()), add = TRUE)

  # provide as.data.frame.class for R < 4.3.0
  as.data.frame.myclass <- function(x, ...) {
    nm <- deparse1(substitute(x))
    as.data.frame(as.character.myclass(x, nm = nm),
                  stringsAsFactors = FALSE)
  }
  assign("as.data.frame.myclass", as.data.frame.myclass, envir = globalenv())
  on.exit(rm("as.data.frame.myclass", envir = globalenv()), add = TRUE)

  obj <- structure(1L, class = "myclass")
  wb <- wb_workbook()$add_worksheet()$add_data(x = obj)

  exp <- "<is><t>myclass: 1</t></is>"
  got <- wb$worksheets[[1]]$sheet_data$cc$is[2]
  expect_equal(exp, got)

  obj <- structure(1.1, class = "myclass")
  wb <- wb_workbook()$add_worksheet()$add_data(x = obj)

  exp <- "<is><t>myclass: 1.1</t></is>"
  got <- wb$worksheets[[1]]$sheet_data$cc$is[2]
  expect_equal(exp, got)

  obj <- structure(TRUE, class = "myclass")
  wb <- wb_workbook()$add_worksheet()$add_data(x = obj)

  exp <- "<is><t>myclass: TRUE</t></is>"
  got <- wb$worksheets[[1]]$sheet_data$cc$is[2]
  expect_equal(exp, got)

})

test_that("date 1904 workbooks to df work", {

  DATE  <- as.Date("2015-02-07") + -10:10
  POSIX <- as.POSIXct("2022-03-02 19:27:35") + -10:10
  time <- data.frame(DATE, POSIX)

  wb <- wb_workbook()
  wb$add_worksheet()$add_data(x = time)
  exp <- wb_to_df(wb)

  wb <- wb_workbook()
  wb$workbook$workbookPr <- '<workbookPr date1904="true"/>'
  wb$add_worksheet()$add_data(x = time)
  got <- wb_to_df(wb)

  expect_equal(exp, got)

  wb <- wb_workbook()
  wb$workbook$workbookPr <- '<workbookPr date1904="1"/>'
  wb$add_worksheet()$add_data(x = time)
  got <- wb_to_df(wb)

  expect_equal(exp, got)

})

test_that("both origins work", {
  expect_equal(
    convert_date(40729),
    convert_date(39267, origin = "1904-01-01")
  )
})

test_that("date 1904 works as expected", {
  set.seed(123)
  vpos <- as.POSIXct(Sys.time() + sample(runif(5) * 1e6, 20, TRUE), tz = "UTC")
  vdte <- as.Date(vpos)
  df <- data.frame(
    int = 1:20,
    dte = vdte,
    pos = vpos,
    var1 = unname(convert_to_excel_date(as.data.frame(vdte), date1904 = TRUE)),
    var2 = unname(convert_to_excel_date(as.data.frame(vpos), date1904 = TRUE))
  )

  wb <- wb_workbook()$add_worksheet()
  wb$workbook$workbookPr <- '<workbookPr date1904="true"/>'
  wb$add_data(x = df)

  got <- wb$to_df(types = c(var1 = "Date", var2 = "POSIXct"))
  expect_equal(df[c("dte", "pos")], got[c("var1", "var2")], ignore_attr = TRUE)

  got <- wb$to_df(convert = FALSE)
  expect_equal(as.character(df$dte), got[["dte"]])
  # ignore rounding differences
  expect_equal(as.Date(df$pos), as.Date(got[["pos"]], tz = "UTC"))
})
