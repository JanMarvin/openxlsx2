test_that("test return values for wb_save", {
  tempFile <- temp_xlsx()
  wb <- wb_add_worksheet(wb_workbook(), "name")
  expect_identical(tempFile, wb_save(wb, tempFile)$path)
  expect_error(wb_save(wb, tempFile), NA)
  expect_error(wb_save(wb, tempFile, overwrite = FALSE))

  expect_identical(tempFile, wb_save(wb, tempFile)$path)
  file.remove(tempFile)
})

# regression test for a typo
test_that("regression test for #248", {

  # Basic data frame
  df <- data.frame(number = 1:3, percent = 4:6 / 100)
  tempFile <- temp_xlsx()

  # no formatting
  expect_silent(write_xlsx(df, tempFile, borders = "columns", overwrite = TRUE))

  # Change column class to percentage
  class(df$percent) <- "percentage"
  expect_silent(write_xlsx(df, tempFile, borders = "columns", overwrite = TRUE))
})


# test for hyperrefs
test_that("creating hyperlinks", {

  # prepare a file
  tempFile <- temp_xlsx()
  sheet <- "test"
  wb <- wb_add_worksheet(wb_workbook(), sheet)
  img <- "D:/somepath/somepicture.png"

  # warning: col and row provided, but not required
  expect_warning(
    linkString <- create_hyperlink(col = 1, row = 4,
                                      text = "test.png", file = img))

  linkString2 <- create_hyperlink(text = "test.png", file = img)

  # col and row not needed
  expect_equal(linkString, linkString2)

  # write file without errors
  write_formula(wb, sheet, x = linkString, startCol = 1, startRow = 1)
  expect_silent(wb_save(wb, tempFile, overwrite = TRUE))

  # TODO: add a check that the written xlsx file contains linkString

})

test_that("write_data2", {
  # create a workbook and add some sheets
  wb <- wb_workbook()

  wb$add_worksheet("sheet1")
  write_data2(wb, "sheet1", mtcars, colNames = TRUE, rowNames = TRUE)

  wb$add_worksheet("sheet2")
  write_data2(wb, "sheet2", cars, colNames = FALSE)

  wb$add_worksheet("sheet3")
  write_data2(wb, "sheet3", letters)

  wb$add_worksheet("sheet4")
  write_data2(wb, "sheet4", as.data.frame(Titanic), startRow = 2, startCol = 2)

  file <- tempfile(fileext = ".xlsx")
  wb_save(wb, file)

  wb1 <- wb_load(file)

  expect_equal(
    mtcars,
    wb_to_df(wb1, "sheet1", rowNames = TRUE),
    ignore_attr = TRUE
  )

  expect_equal(cars, wb_to_df(wb1, "sheet2", colNames = FALSE), ignore_attr = TRUE)

  expect_equal(
    letters,
    as.character(wb_to_df(wb1, "sheet3", colNames = FALSE))
  )

  expect_equal(
    wb_to_df(wb1, "sheet4"),
    as.data.frame(Titanic, stringsAsFactors = FALSE),
    ignore_attr = TRUE
  )

  file.remove(file)
})


test_that("write xlsx", {

  tmp <- temp_xlsx()
  df <- data.frame(a = 1:26, b = letters)

  expect_silent(write_xlsx(df, tmp, tabColour = "#4F81BD"))
  expect_error(write_xlsx(df, tmp, asTable = "YES"))
  expect_error(write_xlsx(df, tmp, sheetName = paste0(letters, letters, collapse = "")))
  expect_error(write_xlsx(df, tmp, zoom = "FULL"))
  expect_silent(write_xlsx(df, tmp, zoom = 200))
  expect_silent(write_xlsx(x = list("S1" = df, "S2" = df), tmp, sheetName = c("Sheet1", "Sheet2")))
  expect_silent(write_xlsx(x = list("S1" = df, "S2" = df), file = tmp))
  expect_silent(write_xlsx(x = list("S1" = df, "S2" = df), tmp, tabColour = "#4F81BD"))
  l <- list(letters)
  names(l) <- paste0(letters, letters, collapse = "")
  expect_warning(write_xlsx(l, tmp))
  expect_error(write_xlsx(df, tmp, gridLines = "YES"))
  expect_silent(write_xlsx(df, tmp, gridLines = FALSE))
  expect_error(write_xlsx(df, tmp, overwrite = FALSE))
  expect_error(write_xlsx(df, tmp, overwrite = "NO"))
  expect_silent(write_xlsx(df, tmp, withFilter = FALSE))
  expect_silent(write_xlsx(df, tmp, withFilter = TRUE))
  expect_error(write_xlsx(df, tmp, withFilter = "NO"))
  expect_silent(write_xlsx(df, tmp, startRow = 2))
  expect_error(write_xlsx(df, tmp, startRow = -1))
  expect_silent(write_xlsx(df, tmp, startCol = "A"))
  expect_silent(write_xlsx(df, tmp, startCol = "2"))
  expect_silent(write_xlsx(df, tmp, startCol = 2))
  expect_error(write_xlsx(df, tmp, startCol = -1))
  expect_error(write_xlsx(df, tmp, col.names = "NO"))
  expect_silent(write_xlsx(df, tmp, col.names = TRUE))
  expect_error(write_xlsx(df, tmp, colNames = "NO"))
  expect_silent(write_xlsx(df, tmp, colNames = TRUE))
  expect_error(write_xlsx(df, tmp, row.names = "NO"))
  expect_silent(write_xlsx(df, tmp, row.names = TRUE))
  expect_error(write_xlsx(df, tmp, rowNames = "NO"))
  expect_silent(write_xlsx(df, tmp, rowNames = TRUE))
  expect_silent(write_xlsx(df, tmp, colWidth = "auto"))
  expect_silent(write_xlsx(list(df, df), tmp, firstActiveCol = 2, firstActiveRow = 2))
  expect_silent(write_xlsx(list(df, df), tmp, firstCol = FALSE, firstRow = FALSE))
  expect_silent(write_xlsx(list(df, df), tmp, firstCol = TRUE, firstRow = TRUE))
  expect_silent(write_xlsx(df, tmp, asTable = TRUE, tableStyle = "TableStyleLight9"))

})


test_that("example", {

  tmp <- temp_xlsx()

  # write to working directory
  expect_silent(write_xlsx(iris, file = tmp, colNames = TRUE))

  expect_silent(
    write_xlsx(iris,
               file = tmp,
               colNames = TRUE
    )
  )

  ## Lists elements are written to individual worksheets, using list names as sheet names if available
  l <- list("IRIS" = iris, "MTCATS" = mtcars, matrix(runif(1000), ncol = 5))
  write_xlsx(l, tmp, colWidths = c(NA, "auto", "auto"))

  expect_silent(write_xlsx(l, tmp,
                           startCol = c(1, 2, 3), startRow = 2,
                           asTable = c(TRUE, TRUE, FALSE), withFilter = c(TRUE, FALSE, FALSE)
  ))

  # specify column widths for multiple sheets
  expect_silent(write_xlsx(l, tmp, colWidths = 20))
  expect_silent(write_xlsx(l, tmp, colWidths = list(100, 200, 300)))
  expect_silent(write_xlsx(l, tmp, colWidths = list(rep(10, 5), rep(8, 11), rep(5, 5))))

})

test_that("writing NA, NaN and Inf", {

  tmp <- temp_xlsx()
  wb <- wb_workbook()

  x <- data.frame(x = c(NA, Inf, -Inf, NaN))
  wb$add_worksheet("Test1")$add_data(x = x)$save(tmp)

  # we wont get the same input back
  exp <- c(NA_character_, "#NUM!", "#NUM!", "#VALUE!")
  got <- unname(unlist(wb_to_df(tmp)))
  expect_equal(exp, got)

  wb$clone_worksheet(old = "Test1", new = "Clone1")$add_data(x = x)$save(tmp)
  got <- unname(unlist(wb_to_df(tmp, "Clone1")))
  expect_equal(exp, got)

  # distinguish between "NA" and NA_character_
  x <- data.frame(x = c(NA, "NA"))
  wb$add_worksheet("Test2")$add_data(x = x)$save(tmp)

  exp <- c(NA_character_, "NA")
  got <- unname(unlist(wb_to_df(tmp, "Test2")))
  expect_equal(exp, got)

  wb$clone_worksheet(old = "Test2", new = "Clone2")$add_data(x = x)$save(tmp)
  got <- unname(unlist(wb_to_df(tmp, "Clone2")))
  expect_equal(exp, got)

})


test_that("writing NA, NaN and Inf", {

  tmp <- temp_xlsx()
  wb <- wb_workbook()

  x <- data.frame(x = c(NA, Inf, -Inf, NaN))
  wb$add_worksheet("Test1")$add_data(x = x, na.strings = NULL)$save(tmp)
  wb$add_worksheet("Test2")$add_data_table(x = x, na.strings = "N/A")$save(tmp)
  wb$add_worksheet("Test3")$add_data(x = x, na.strings = "N/A")$save(tmp)

  exp <- c(NA, "s", "s", "s")
  got <- unname(unlist(attr(wb_to_df(tmp, "Test1"), "tt")))
  expect_equal(exp, got)

  exp <- c("N/A", "#NUM!", "#NUM!", "#VALUE!")
  got <- unname(unlist(wb_to_df(tmp, "Test2")))
  expect_equal(exp, got)

  wb$clone_worksheet("Test1", "Clone1")$add_data(x = x, na.strings = NULL)$save(tmp)
  wb$clone_worksheet("Test3", "Clone3")$add_data(x = x, na.strings = "N/A")$save(tmp)

  exp <- c(NA, "s", "s", "s")
  got <- unname(unlist(attr(wb_to_df(tmp, "Test1"), "tt")))
  expect_equal(exp, got)

  exp <- c("N/A", "#NUM!", "#NUM!", "#VALUE!")
  got <- unname(unlist(wb_to_df(tmp, "Test2")))
  expect_equal(exp, got)

})


test_that("write cells without data", {

  temp <- temp_xlsx()
  tmp <- temp_dir()

  dat <- as.data.frame(matrix(NA, 2, 2))
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = dat, startRow = 2, startCol = 2, na.strings = NULL, colNames = FALSE)

  wb$worksheets[[1]]$sheet_data$cc$c_t <- ""

  # # created an empty canvas that can be styled
  # wb$add_fill(dims = "B2:C3", color = wb_colour(hex = "FFFFFF00"))
  # wb$add_border(dims = "B2:C3")

  wb$save(temp)

  unzip(temp, exdir = tmp)

  exp <- structure(
    list(
      r = c("B2", "C2", "B3", "C3"),
      row_r = c("2", "2", "3", "3"),
      c_r = c("B", "C", "B", "C"),
      c_s = c("", "", "", ""),
      c_t = c("", "", "", ""),
      c_cm = c("", "", "", ""),
      c_ph = c("", "", "", ""),
      c_vm = c("", "", "", ""),
      v = c("", "", "", ""),
      f = c("", "", "", ""),
      f_t = c("", "", "", ""),
      f_ref = c("", "", "", ""),
      f_ca = c("", "", "", ""),
      f_si = c("", "", "", ""),
      is = c("", "", "", ""),
      typ = c("3", "3", "3", "3")
    ),
    row.names = c(NA, 4L),
    class = "data.frame")
  got <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(exp, got)

  sheet <- paste0(tmp, "/xl/worksheets/sheet1.xml")
  exp <- "<sheetData><row r=\"2\"><c r=\"B2\"/><c r=\"C2\"/></row><row r=\"3\"><c r=\"B3\"/><c r=\"C3\"/></row></sheetData>"
  got <- xml_node(sheet, "worksheet", "sheetData")
  expect_equal(exp, got)

})

test_that("write_xlsx with na.strings", {

  df <- data.frame(
    num = c(1, -99, 3, NA_real_),
    char = c("hello", "99", "3", NA_character_)
  )

  test <- temp_xlsx()
  write_xlsx(df, file = test)

  exp <- df
  got <- read_xlsx(test)
  expect_equal(exp, got, ignore_attr = TRUE)

  write_xlsx(df, file = test, na.strings = "N/A")
  got <- read_xlsx(test, na.strings = "N/A")
  expect_equal(exp, got, ignore_attr = TRUE)

  exp$num[exp$num == -99] <- NA
  got <- read_xlsx(test, na.strings = "N/A", na.numbers = -99)
  expect_equal(exp, got, ignore_attr = TRUE)

})

test_that("write & load file with chartsheet", {

  fl <- system.file("extdata", "mtcars_chart.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)

  wb$worksheets[[1]]$sheetPr <- xml_node_create(
    "sheetPr",
    xml_children = xml_node_create(
      "tabColor",
      xml_attributes = c(rgb = "FF00FF00")))

  temp <- temp_xlsx()
  expect_silent(wb$save(temp))
  expect_silent(wb2 <- wb_load(temp))

})

test_that("escaping of inlinestrings works", {

  temp <- temp_xlsx()
  wb <- wb_workbook()$
    add_worksheet("Test")$
    add_data(dims = "A1", x = "A & B")$
    save(temp)

  exp <- "A & B"
  got <- wb_to_df(wb, colNames = FALSE)$A
  expect_equal(exp, got)

  got <- wb_to_df(temp, colNames = FALSE)$A
  expect_equal(exp, got)

  wb2 <- wb_load(temp)
  got <- wb_to_df(wb2, colNames = FALSE)$A
  expect_equal(exp, got)

})
