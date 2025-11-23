
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
  expect_equal(attr(z, "tt")$A, 5L)

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
  vpos <- .POSIXct(Sys.time() + sample(runif(5) * 1e6, 20, TRUE), tz = "UTC")
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

test_that("date conversion works", {

  wb <- wb_workbook()$add_worksheet()
  wb$add_data(dims = "A1:D1", x = as.Date(paste0("2025-0", 1:4, "-01")), col_names = FALSE)
  wb$add_data(dims = "A2:D4", x = matrix(1:12, 3, 4), col_names = FALSE)

  # column name is converted date, column is numeric
  df <- wb$to_df()
  expect_true(is.numeric(df$`2025-01-01`))

  # column name is converted date, column is character
  df <- wb$to_df(convert = FALSE)
  expect_true(is.character(df$`2025-01-01`))

  # column name is spreadsheet date, column is numeric
  df <- wb$to_df(convert = TRUE, detect_dates = FALSE)
  expect_true(is.numeric(df$`45658`))

  # column name is spreadsheet date, column is character
  df <- wb$to_df(convert = FALSE, detect_dates = FALSE)
  expect_true(is.character(df$`45658`))

  # conversion works for rownames
  df <- data.frame(
    x = as.Date(paste0("2025-0", 1:4, "-01")),
    y = 1:4
  )
  wb <- wb_workbook()$add_worksheet()
  wb$add_data(x = df)
  df <- wb$to_df(row_names = TRUE)
  exp <- c("2025-01-01", "2025-02-01", "2025-03-01", "2025-04-01")
  got <- rownames(df)
  expect_equal(exp, got)
})

test_that("conversion works", {

  wb <- wb_workbook()$add_worksheet()
  # row 1 column name
  wb$add_data(dims = "A1", x = "Var1")
  wb$add_data(dims = "B1", x = "Var2")
  wb$add_data(dims = "C1", x = "Var3")
  # row 2 character
  wb$add_data(dims = "A2", x = "2024-01-31")
  wb$add_data(dims = "B2", x = "2024-01-31")
  wb$add_data(dims = "C2", x = "2024-01-31")
  # various dates
  wb$add_data(dims = "A3", x = as.Date("2024-02-01"))
  wb$add_data(dims = "B3", x = structure(1234, units = "secs", class = c("hms", "difftime")))
  wb$add_data(dims = "C3", x = as.POSIXct("2024-04-25 08:47:03", tz = "UTC"))

  exp <- structure(
    list(
      Var1 = c("2024-01-31", "2024-02-01"),
      Var2 = c("2024-01-31", "00:20:34"),
      Var3 = c("2024-01-31", "2024-04-25 08:47:03")
    ), row.names = 2:3, class = "data.frame")
  got <- wb$to_df()
  expect_equal(exp, got)

  got <- wb$to_df(convert = FALSE)
  expect_equal(exp, got)

  exp <- structure(
    list(
      `2024-01-31` = structure(19754, class = "Date"),
      `2024-01-31` = "00:20:34",
      `2024-01-31` = structure(1714034823, class = c("POSIXct", "POSIXt"), tzone = "UTC")
    ), row.names = 3L, class = "data.frame")

  got <- wb$to_df(start_row = 2)
  expect_equal(exp, got)

  exp <- structure(
    list(
      `2024-01-31` = "2024-02-01",
      `2024-01-31` = "00:20:34",
      `2024-01-31` = "2024-04-25 08:47:03"
    ), row.names = 3L, class = "data.frame")
  got <- wb$to_df(start_row = 2, convert = FALSE)
  expect_equal(exp, got)

  exp <- structure(c(NA, 19838), class = "Date")
  expect_warning(got <- wb$to_df(types = c(Var3 = "Date")), "coercion")
  expect_equal(exp, got$Var3)

  exp <- structure(c(NA, 1706745600), class = c("POSIXct", "POSIXt"), tzone = "UTC")
  expect_warning(got <- wb$to_df(types = c(Var1 = "POSIXct")), "coercion")
  expect_equal(exp, got$Var1)
})

test_that("Ignoring 'GENERAL' numfmt works", {

  # inspired by https://stackoverflow.com/q/25158969
  df <- data.frame(
    date = Sys.Date() - seq_len(5),
    num = seq_len(5)
  )

  wb <- wb_workbook()$add_worksheet()$
    add_data(x = df)$
    add_numfmt(dims = "B2:B6", numfmt = "GENERAL")

  got <- wb$to_df()
  expect_equal(df, got, ignore_attr = TRUE)

})
