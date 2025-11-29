test_that("test return values for wb_save", {
  tempFile <- temp_xlsx()
  wb <- wb_add_worksheet(wb_workbook(), "name")
  expect_identical(tempFile, wb_save(wb, tempFile)$path)
  expect_error(wb_save(wb, tempFile), NA)
  expect_error(wb_save(wb, tempFile, overwrite = FALSE))

  expect_identical(tempFile, wb_save(wb, tempFile)$path)
})

# regression test for a typo
test_that("regression test for #248", {

  # Basic data frame
  df <- data.frame(number = 1:3, percent = 4:6 / 100)
  tempFile <- temp_xlsx()

  # no formatting
  expect_silent(write_xlsx(df, tempFile, overwrite = TRUE))

  # Change column class to percentage
  class(df$percent) <- "percentage"
  expect_silent(write_xlsx(df, tempFile, overwrite = TRUE))
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
  wb$add_formula(sheet, x = linkString, start_col = 1, start_row = 1)
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

  file <- temp_xlsx()
  wb_save(wb, file)

  wb1 <- wb_load(file)

  expect_equal(
    mtcars,
    wb_to_df(wb1, "sheet1", row_names = TRUE),
    ignore_attr = TRUE
  )

  expect_equal(cars, wb_to_df(wb1, "sheet2", col_names = FALSE), ignore_attr = TRUE)

  expect_equal(
    letters,
    as.character(wb_to_df(wb1, "sheet3", col_names = FALSE))
  )

  expect_equal(
    wb_to_df(wb1, "sheet4"),
    as.data.frame(Titanic, stringsAsFactors = FALSE),
    ignore_attr = TRUE
  )

})


test_that("write xlsx", {

  tmp <- temp_xlsx()
  df <- data.frame(a = 1:26, b = letters)

  ## FIXME these are not actually tests. They simply check that no error is printed
  expect_silent(write_xlsx(df, tmp, tab_colour = "#4F81BD"))
  expect_error(write_xlsx(df, tmp, as_table = "YES"))
  expect_error(write_xlsx(df, tmp, sheet_name = paste0(letters, letters, collapse = "")))
  expect_error(write_xlsx(df, tmp, zoom = "FULL"))
  expect_silent(write_xlsx(df, tmp, zoom = 200))
  expect_silent(write_xlsx(x = list("S1" = df, "S2" = df), tmp, sheet_name = c("Sheet1", "Sheet2")))
  expect_silent(write_xlsx(x = list("S1" = df, "S2" = df), file = tmp))
  expect_silent(write_xlsx(x = list("S1" = df, "S2" = df), tmp, tab_colour = "#4F81BD"))
  l <- list(letters)
  names(l) <- paste0(letters, letters, collapse = "")
  expect_warning(write_xlsx(l, tmp))
  expect_error(write_xlsx(df, tmp, grid_lines = "YES"))
  expect_silent(write_xlsx(df, tmp, grid_lines = FALSE))
  expect_error(write_xlsx(df, tmp, overwrite = FALSE))
  expect_error(write_xlsx(df, tmp, overwrite = "NO"))
  expect_silent(write_xlsx(df, tmp, with_filter = FALSE))
  expect_silent(write_xlsx(df, tmp, with_filter = TRUE))
  expect_error(write_xlsx(df, tmp, with_filter = "NO"))
  expect_silent(write_xlsx(df, tmp, start_row = 2))
  expect_error(write_xlsx(df, tmp, start_row = -1))
  expect_silent(write_xlsx(df, tmp, start_col = "A"))
  expect_silent(write_xlsx(df, tmp, start_col = "2"))
  expect_silent(write_xlsx(df, tmp, start_col = 2))
  expect_error(write_xlsx(df, tmp, start_col = -1))
  expect_error(write_xlsx(df, tmp, col.names = "NO"))
  expect_silent(write_xlsx(df, tmp, col.names = TRUE))
  expect_error(write_xlsx(df, tmp, col_names = "NO"))
  expect_silent(write_xlsx(df, tmp, col_names = TRUE))
  expect_error(write_xlsx(df, tmp, row.names = "NO"))
  expect_silent(write_xlsx(df, tmp, row.names = TRUE))
  expect_error(write_xlsx(df, tmp, row_names = "NO"))
  expect_silent(write_xlsx(df, tmp, row_names = TRUE))
  expect_silent(write_xlsx(df, tmp, col_widths = "auto"))
  expect_silent(write_xlsx(list(df, df), tmp, first_active_col = 2, first_active_row = 2))
  expect_silent(write_xlsx(list(df, df), tmp, first_col = FALSE, first_row = FALSE))
  expect_silent(write_xlsx(list(df, df), tmp, first_col = TRUE, first_row = TRUE))
  expect_silent(write_xlsx(df, tmp, as_table = TRUE, table_style = "TableStyleLight9"))

})


test_that("example", {

  tmp <- temp_xlsx()

  # write to working directory
  expect_silent(write_xlsx(iris, file = tmp, col_names = TRUE))

  expect_silent(
    write_xlsx(iris,
               file = tmp,
               col_names = TRUE
    )
  )

  ## Lists elements are written to individual worksheets, using list names as sheet names if available
  l <- list("IRIS" = iris, "MTCATS" = mtcars, matrix(runif(1000), ncol = 5))
  write_xlsx(l, tmp, col_widths = c(NA, "auto", "auto"))

  expect_silent(write_xlsx(l, tmp,
                           start_col = c(1, 2, 3), start_row = 2,
                           as_table = c(TRUE, TRUE, FALSE), with_filter = c(TRUE, FALSE, FALSE)
  ))

  # specify column widths for multiple sheets
  expect_silent(write_xlsx(l, tmp, col_widths = 20))
  expect_silent(write_xlsx(l, tmp, col_widths = list(100, 200, 300)))
  expect_silent(write_xlsx(l, tmp, col_widths = list(rep(10, 5), rep(8, 11), rep(5, 5))))

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
  x <- data.frame(x = c(NA, "NA"), stringsAsFactors = FALSE)
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

  exp <- c(NA, 0L, 0L, 0L)
  got <- unname(unlist(attr(wb_to_df(tmp, "Test1", keep_attributes = TRUE), "tt")))
  expect_equal(exp, got)

  exp <- c("N/A", "#NUM!", "#NUM!", "#VALUE!")
  got <- unname(unlist(wb_to_df(tmp, "Test2", keep_attributes = TRUE)))
  expect_equal(exp, got)

  wb$clone_worksheet("Test1", "Clone1")$add_data(x = x, na.strings = NULL)$save(tmp)
  wb$clone_worksheet("Test3", "Clone3")$add_data(x = x, na.strings = "N/A")$save(tmp)

  exp <- c(NA, 0L, 0L, 0L)
  got <- unname(unlist(attr(wb_to_df(tmp, "Test1", keep_attributes = TRUE), "tt")))
  expect_equal(exp, got)

  exp <- c("N/A", "#NUM!", "#NUM!", "#VALUE!")
  got <- unname(unlist(wb_to_df(tmp, "Test2")))
  expect_equal(exp, got)

})


test_that("write cells without data", {

  temp <- temp_xlsx()
  tmp_dir <- temp_dir()

  dat <- as.data.frame(matrix(NA, 2, 2))
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = dat, start_row = 2, start_col = 2, na.strings = NULL, col_names = FALSE)

  wb$worksheets[[1]]$sheet_data$cc$c_t <- ""

  # # created an empty canvas that can be styled
  # wb$add_fill(dims = "B2:C3", color = wb_colour(hex = "FFFFFF00"))
  # wb$add_border(dims = "B2:C3")

  wb$save(temp)

  unzip(temp, exdir = tmp_dir)

  exp <- structure(
    list(
      r = c("B2", "C2", "B3", "C3"),
      row_r = c("2", "2", "3", "3"),
      c_r = c("B", "C", "B", "C"),
      c_s = c("", "", "", ""),
      c_t = c("", "", "", ""),
      v = c("", "", "", ""),
      f = c("", "", "", ""),
      f_attr = c("", "", "", ""),
      is = c("", "", "", "")
    ),
    row.names = c(NA, 4L),
    class = "data.frame")
  got <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(exp, got)

  sheet <- paste0(tmp_dir, "/xl/worksheets/sheet1.xml")
  exp <- "<sheetData><row r=\"2\"/><row r=\"3\"/></sheetData>"
  got <- xml_node(sheet, "worksheet", "sheetData")
  expect_equal(exp, got)

})

test_that("write_xlsx with na.strings", {

  df <- data.frame(
    num = c(1, -99, 3, NA_real_),
    char = c("hello", "99", "3", NA_character_),
    stringsAsFactors = FALSE
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

  fl <- testfile_path("mtcars_chart.xlsx")
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
  got <- wb_to_df(wb, col_names = FALSE)$A
  expect_equal(exp, got)

  got <- wb_to_df(temp, col_names = FALSE)$A
  expect_equal(exp, got)

  wb2 <- wb_load(temp)
  got <- wb_to_df(wb2, col_names = FALSE)$A
  expect_equal(exp, got)

})

test_that("write_xlsx() works", {

  tmp <- temp_xlsx()
  write_xlsx(mtcars, tmp, sheet_name = "test")
  exp <- c(test = "test")
  got <- wb_load(tmp)$get_sheet_names()
  expect_equal(exp, got)

})

test_that("write_xlsx() freezing rows works", {

  tmp <- temp_xlsx()

  wb <- write_xlsx(list(mtcars, mtcars), tmp, first_row = TRUE, first_col = TRUE, tab_color = wb_color("green"))

  # tabColor
  exp <- c(
    "<sheetPr><tabColor rgb=\"FF00FF00\"/></sheetPr>",
    "<sheetPr><tabColor rgb=\"FF00FF00\"/></sheetPr>"
  )
  got <- c(
    wb$worksheets[[1]]$sheetPr,
    wb$worksheets[[2]]$sheetPr
  )
  expect_equal(exp, got)

  # firstCol/firstRow
  exp <- c(
    "<pane ySplit=\"1\" xSplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>",
    "<pane ySplit=\"1\" xSplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  )
  got <- c(
    wb$worksheets[[1]]$freezePane,
    wb$worksheets[[2]]$freezePane
  )
  expect_equal(exp, got)

  wb <- write_xlsx(list(mtcars, mtcars), tmp, first_active_row = 4, first_active_col = 3)

  # firstActiveCol/firstActiveRow
  exp <- c(
    "<pane ySplit=\"3\" xSplit=\"2\" topLeftCell=\"C4\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>",
    "<pane ySplit=\"3\" xSplit=\"2\" topLeftCell=\"C4\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  )
  got <- c(
    wb$worksheets[[1]]$freezePane,
    wb$worksheets[[2]]$freezePane
  )
  expect_equal(exp, got)

})

test_that("write_xlsx works with colour", {

  tmp <- temp_xlsx()

  wb <- write_xlsx(mtcars, tmp, tab_colour = "green")

  exp <- "<sheetPr><tabColor rgb=\"FF00FF00\"/></sheetPr>"
  got <- wb$worksheets[[1]]$sheetPr
  expect_equal(exp, got)

})

test_that("write_xlsx with base font settings", {
  tmp <- temp_xlsx()
  df <- data.frame(a = 1:5, b = letters[1:5])

  # Test with font size
  expect_silent(write_xlsx(df, tmp, font_size = 14))

  # Test with font color
  expect_silent(write_xlsx(df, tmp, font_color = wb_color(theme = "2")))

  # Test with font name
  expect_silent(write_xlsx(df, tmp, font_name = "Arial"))

  # Test with all font parameters
  expect_silent(write_xlsx(df, tmp,
    font_size = 12,
    font_color = wb_color(theme = "1"),
    font_name = "Calibri"
  ))

  # Load and verify font settings
  wb <- write_xlsx(df, tmp,
    font_size = 16,
    font_color = wb_color(auto = TRUE),
    font_name = "Times New Roman"
  )
  font <- wb_get_base_font(wb)
  expect_equal(font$size$val, "16")
  expect_equal(font$name$val, "Times New Roman")
  expect_equal(font$color$auto, "1")
})

test_that("glue is supported", {
  x <- structure("foo", class = c("glue", "character"))

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = x)

  expect_equal("foo", wb_to_df(wb, col_names = FALSE)$A)
})

test_that("table names work in write_xlsx", {
  wb <- write_xlsx(list(data = cars), as_table = TRUE, table_style = "TableStyleLight18", table_name = "tblcars")

  exp <- "tblcars"
  got <- wb$get_named_regions(tables = TRUE)$name
  expect_equal(exp, got)


  wb <- write_xlsx(list(cars = cars, mtcars = mtcars), as_table = TRUE, table_style = "TableStyleLight18", table_name = "tblcars")

  exp <- c("tblcars", "tblcars1")
  got <- wb$get_named_regions(tables = TRUE)$name
  expect_equal(exp, got)


  wb <- write_xlsx(list(cars = cars, mtcars = mtcars), as_table = TRUE, table_style = "TableStyleLight18", table_name = c("tblcars", "tblmtcars"))

  exp <- c("tblcars", "tblmtcars")
  got <- wb$get_named_regions(tables = TRUE)$name
  expect_equal(exp, got)
})
