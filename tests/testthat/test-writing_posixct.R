test_that("Writing Posixct with write_data & write_datatable", {
  op <- options("openxlsx2.datetimeFormat" = "dd/mm/yy hh:mm")
  on.exit(options(op), add = TRUE)

  tstart <- strptime("30/05/2017 08:30", "%d/%m/%Y %H:%M", tz = "CET")
  TimeDT <- c(0, 5, 10, 15, 30, 60, 120, 180, 240, 480, 720, 1440) * 60 + tstart
  df <- data.frame(TimeDT, TimeTxt = format(TimeDT, "%Y-%m-%d %H:%M"))

  wb <- wb_workbook()
  wb$add_worksheet("write_data")
  wb$add_worksheet("write_datatable")

  wb$add_data("write_data", df, start_col = 2, start_row = 3, row_names = FALSE)
  wb$add_data_table("write_datatable", df, start_col = 2, start_row = 3)

  # wb_open(wb)

  wd <- wb$worksheets[[1]]$sheet_data$cc
  wdt <- wb$worksheets[[2]]$sheet_data$cc

  expect_equal(wd$v[[3]], "42885.3541666667")
  expect_equal(wd$is[[4]], "<is><t>2017-05-30 08:30</t></is>")

})

# missing datetime is not yet implemented
test_that("Writing mixed EDT/EST Posixct with write_data & write_datatable", {
  op <- options("openxlsx2.datetimeFormat" = "dd/mm/yy hh:mm")
  on.exit(options(op), add = TRUE)

  tstart1 <- as.POSIXct("12/03/2018 08:30", format = "%d/%m/%Y %H:%M")
  tstart2 <- as.POSIXct("10/03/2018 08:30", format = "%d/%m/%Y %H:%M")
  TimeDT1 <- c(NA, 0, 10, 30, 60, 120, 240, 720, 1440) * 60 + tstart1
  TimeDT2 <- c(0, 10, 30, 60, 120, 240, 720, NA, 1440) * 60 + tstart2

  df <- data.frame(
    timeval = c(TimeDT1, TimeDT2),
    timetxt = format(c(TimeDT1, TimeDT2), "%Y-%m-%d %H:%M")
  )

  wb <- wb_workbook()
  wb$add_worksheet("write_data")
  wb$add_worksheet("write_datatable")

  wb$add_data("write_data", df, start_col = 2, start_row = 3, row_names = FALSE)
  wb$add_data_table("write_datatable", df, start_col = 2, start_row = 3)

  xlsxFile <- temp_xlsx()
  wb_save(wb, xlsxFile, TRUE)

  wb_s1 <- wb_to_df(xlsxFile, sheet = "write_data")
  wb_s2 <- wb_to_df(xlsxFile, sheet = "write_datatable")

  # compare sheet 1
  exp <- df$timeval
  got <- wb_s1$timeval
  expect_equal(exp, got, tolerance = 10 ^ -10, ignore_attr = "tzone")

  exp <- df$timetxt
  got <- wb_s1$timetxt
  expect_equal(exp, got)

  # compare sheet 2
  exp <- df$timeval
  got <- wb_s2$timeval
  expect_equal(exp, got, tolerance = 10 ^ -10, ignore_attr = "tzone")

  exp <- df$timetxt
  got <- wb_s2$timetxt
  expect_equal(exp, got)

})

test_that("numfmt escaping works", {

  op <- options(
    "openxlsx2.datetimeFormat" = "yyyy\\/mm\\/dd",
    "openxlsx2.dateFormat" = "mm/dd/yyyy"
  )
  on.exit(options(op), add = TRUE)

  test_data <- data.frame(
    datetime_col = as.POSIXct("2023-12-31 00:00:00"),
    date_col = as.Date("2023-12-31")
  )
  wb <- wb_workbook()$add_worksheet()$add_data(x = test_data)

  exp <- c("<numFmt numFmtId=\"165\" formatCode=\"mm\\/dd\\/yyyy\"/>",
          "<numFmt numFmtId=\"166\" formatCode=\"yyyy\\/mm\\/dd\"/>")
  got <- wb$styles_mgr$styles$numFmts
  expect_equal(exp, got)

})
