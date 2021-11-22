
test_that("convertToDate() works", {
  dates <- as.Date("2015-02-07") + -10:10
  origin <- 25569L
  n <- as.integer(dates) + origin

  expect_identical(convertToDate(n), dates)

  earlyDate <- as.Date("1900-01-03")
  serialDate <- 3
  expect_identical(convertToDate(serialDate), earlyDate)
})


test_that("convertToDateTime() works", {
  x <- 43037 + 2 / 1440
  res <- convertToDateTime(x, tz = Sys.timezone())
  exp <- as.POSIXct("2017-10-29 00:02:00", tz = Sys.timezone())
  expect_identical(res, exp)

  x <- 43037 + 2 / 1440 + 1 / 86400
  res <- convertToDateTime(x, tz = Sys.timezone())
  exp <- as.POSIXct("2017-10-29 00:02:01", tz = Sys.timezone())
  expect_identical(res, exp)

  x <- 43037 + 2.50 / 1440
  res <- convertToDateTime(x, tz = Sys.timezone())
  exp <- as.POSIXct("2017-10-29 00:02:30", tz = Sys.timezone())
  expect_identical(res, exp)

  # Not sure what this is suppose to be
  # x <- 43037 + 2 / 1440 + 12.12 / 86400
  # x_datetime <- convertToDateTime(x, tz = "UTC")
  # attr(x_datetime, "tzone") <- "UTC"
})
