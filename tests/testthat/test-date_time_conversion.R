
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
  res <- convert_datetime(x, tx = Sys.timezone())
  exp <- as.POSIXct("2017-10-29 00:02:00", tz = Sys.timezone())
  expect_equal(res, exp, ignore_attr = "tzone")

  x <- 43037 + 2 / 1440 + 1 / 86400
  res <- convert_datetime(x, tx = Sys.timezone())
  exp <- as.POSIXct("2017-10-29 00:02:01", tz = Sys.timezone())
  expect_equal(res, exp, ignore_attr = "tzone")

  x <- 43037 + 2.50 / 1440
  res <- convert_datetime(x, tx = Sys.timezone())
  exp <- as.POSIXct("2017-10-29 00:02:30", tz = Sys.timezone())
  expect_equal(res, exp, ignore_attr = "tzone")

  x <- 43037 + 2 / 1440 + 12.12 / 86400
  x_datetime <- convert_datetime(x, tx = "UTC")
  attr(x_datetime, "tzone") <- "UTC"
})
