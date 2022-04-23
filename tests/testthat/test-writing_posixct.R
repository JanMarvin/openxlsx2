test_that("Writing Posixct with writeData & writeDataTable", {
  options("openxlsx2.datetimeFormat" = "dd/mm/yy hh:mm")

  tstart <- strptime("30/05/2017 08:30", "%d/%m/%Y %H:%M", tz = "CET")
  TimeDT <- c(0, 5, 10, 15, 30, 60, 120, 180, 240, 480, 720, 1440) * 60 + tstart
  df <- data.frame(TimeDT, TimeTxt = format(TimeDT, "%Y-%m-%d %H:%M"))

  wb <- wb_workbook()
  wb$add_worksheet("writeData")
  wb$add_worksheet("writeDataTable")

  writeData(wb, "writeData", df, startCol = 2, startRow = 3, rowNames = FALSE)
  writeDataTable(wb, "writeDataTable", df, startCol = 2, startRow = 3)

  # wb_open(wb)

  wd <- wb$worksheets[[1]]$sheet_data$cc
  wdt <- wb$worksheets[[2]]$sheet_data$cc

  expect_equal(wd$v[[3]], "42885.3541666667")
  expect_equal(wd$is[[4]], "<is><t>2017-05-30 08:30</t></is>")

  options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
})

# missing datetime is not yet implemented
test_that("Writing mixed EDT/EST Posixct with writeData & writeDataTable", {
  options("openxlsx2.datetimeFormat" = "dd/mm/yy hh:mm")

  tstart1 <- as.POSIXct("12/03/2018 08:30", format = "%d/%m/%Y %H:%M")
  tstart2 <- as.POSIXct("10/03/2018 08:30", format = "%d/%m/%Y %H:%M")
  TimeDT1 <- c(NA, 0, 10, 30, 60, 120, 240, 720, 1440) * 60 + tstart1
  TimeDT2 <- c(0, 10, 30, 60, 120, 240, 720, NA, 1440) * 60 + tstart2

  df <- data.frame(
    timeval = c(TimeDT1, TimeDT2),
    timetxt = format(c(TimeDT1, TimeDT2), "%Y-%m-%d %H:%M")
  )

  wb <- wb_workbook()
  wb$add_worksheet("writeData")
  wb$add_worksheet("writeDataTable")

  writeData(wb, "writeData", df, startCol = 2, startRow = 3, rowNames = FALSE)
  writeDataTable(wb, "writeDataTable", df, startCol = 2, startRow = 3)

  # xlsx file is broken‚ <NA> where some missing value is expected.
  # TODO check: looks alright in LibreOffice
  # wb_open(wb)
  xlsxFile <- temp_xlsx()
  wb_save(wb, xlsxFile, TRUE)

  wb_s1 <- wb_to_df(xlsxFile, sheet = "writeData")
  wb_s2 <- wb_to_df(xlsxFile, sheet = "writeDataTable")

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

  options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
})
