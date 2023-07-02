test_that("wb_to_df", {

  ###########################################################################
  # numerics, dates, missings, bool and string
  xlsxFile <- testfile_path("readTest.xlsx")
  expect_silent(wb1 <- wb_load(xlsxFile))

  # import workbook
  exp <- structure(
    list(c(TRUE, TRUE, TRUE, FALSE, FALSE, FALSE, NA, FALSE, FALSE, NA),
         c(1, NA, 2, 2, 3, 1, NA, 2, 3, 1),
         rep(NA_real_, 10),
         c("1", "#NUM!", "1.34", NA, "1.56", "1.7", NA, "23", "67.3", "123"),
         c("a", "b", "c", "#NUM!", "e", "f", NA, "h", "i", NA),
         structure(
           c(16473, 16472, 16471, NA, NA, 16468, 16467, 16466, 16465, 16464),
           class = "Date"
         ),
         c("3209324 This", NA, NA, NA, NA, NA, NA, NA, NA, NA),
         c("#DIV/0!", NA, "#NUM!", NA, NA, NA, NA, NA, NA, NA)),
    .Names = c("Var1", "Var2", NA, "Var3", "Var4", "Var5", "Var6", "Var7"),
    row.names = c(2:11),
    class = "data.frame"
  )
  got <- wb_to_df(wb1)
  expect_equal(exp, got, ignore_attr = TRUE)

  # do not convert first row to colNames
  got <- wb_to_df(wb1, colNames = FALSE)
  expect_equal(int2col(seq_along(got)), names(got))

  # do not try to identify dates in the data
  got <- wb_to_df(wb1, detectDates = FALSE)
  expect_equal(convert_to_excel_date(df = exp["Var5"], date1904 = FALSE),
               got["Var5"])

  expect_warning(convertToExcelDate(df = exp["Var5"], date1904 = FALSE), "deprecated")

  # return the underlying Excel formula instead of their values
  got <- wb_to_df(wb1, showFormula = TRUE)
  expect_equal("1/0", got$Var7[1])

  # read dimension withot colNames
  got <- wb_to_df(wb1, dims = "A2:C5", colNames = FALSE)
  test <- data.frame(A = c(TRUE, TRUE, TRUE, FALSE),
                     B = c(1, NA, 2, 2),
                     C = rep(NA_real_, 4))
  expect_equal(test, got, ignore_attr = TRUE)

  # read selected cols
  got <- wb_to_df(wb1, cols = c(1:2, 7))
  expect_equal(exp[c(1, 2, 7)], got, ignore_attr = TRUE)

  # read selected cols
  got <- wb_to_df(wb1, cols = c("A", "B", "G"))
  expect_equal(exp[c(1, 2, 7)], got, ignore_attr = TRUE)

  # read selected rows
  got <- wb_to_df(wb1, rows = c(1, 4, 6))
  got[c(4, 7)] <- lapply(got[c(4, 7)], as.character)
  expect_equal(exp[c(3, 5), ], got, ignore_attr = TRUE)

  # convert characters to numerics and date (logical too?)
  got <- wb_to_df(wb1, convert = FALSE)
  chrs <- exp
  chrs[seq_along(chrs)] <- lapply(chrs[seq_along(chrs)], as.character)
  expect_equal(chrs, got, ignore_attr = TRUE)

  # # erase empty Rows from dataset
  # not gonna test this :) just want to mention how blazing fast it is now.
  # got <- wb_to_df(wb1, sheet = 3, skipEmptyRows = TRUE)

  # erase rmpty Cols from dataset
  got <- wb_to_df(wb1, skipEmptyCols = TRUE)
  expect_equal(exp[c(1, 2, 4, 5, 6, 7, 8)], got, ignore_attr = TRUE)

  # # convert first row to rownames
  # wb_to_df(wb1, sheet = 3, dims = "C6:G9", rowNames = TRUE)

  # define type of the data.frame
  got <- wb_to_df(wb1, cols = c(1, 4), types = c("Var1" = 0, "Var3" = 1))
  test <- exp[c("Var1", "Var3")]
  test["Var1"] <- lapply(test["Var1"], as.character)
  suppressWarnings(test["Var3"] <- lapply(test["Var3"], function(x) as.numeric(replace(x, x == "#NUM!", "NaN"))))
  expect_equal(test, got, ignore_attr = TRUE)

  # start in row 5
  got <- wb_to_df(wb1, startRow = 5, colNames = FALSE)
  test <- exp[4:10, ]
  names(test) <- int2col(seq_along(test))
  test[c("D", "G", "H")] <- lapply(test[c("D", "G", "H")], as.numeric)
  expect_equal(test, got, ignore_attr = TRUE)

  # na string
  got <- wb_to_df(wb1, na.strings = "")
  expect_equal("#N/A", got$Var7[2], ignore_attr = TRUE)


  ###########################################################################
  # inlinestr
  xlsxFile <- testfile_path("inline_str.xlsx")
  expect_silent(wb2 <- wb_load(xlsxFile))

  exp <- data.frame(
    PairIndex = c(rep(1, 8), 2, 2), Drug1 = "abc", Drug2 = "def", Conc1 = 10000,
    Conc2 = c(10000, 3000, 1000, 300, 100, 30, 10, 0, 0, 10),
    Response = c(
      -1.79607109448082, 1.01028999064546, 0.449017773620206, 0,
      0.898035547240412, 0.112254443405051, 3.14312441534144,
      1.45930776426567, 1.68381665107577, -0.78578110383536),
    concUnit = "nM",
    stringsAsFactors = FALSE)
  rownames(exp) <- seq(2, nrow(exp) + 1)
  # read dataset with inlinestr
  got <- wb_to_df(wb2)
  expect_equal(exp, got, ignore_attr = TRUE)


  ###########################################################################
  # named_region // namedRegion
  xlsxFile <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
  expect_silent(wb3 <- wb_load(xlsxFile))

  # read dataset with named_region (returns global first)
  exp <- data.frame(A = "S2A1", B = "S2B1")
  got <- wb_to_df(wb3, named_region = "MyRange", colNames = FALSE)
  expect_equal(exp, got, ignore_attr = TRUE)

  # read named_region from sheet
  exp <- data.frame(A = "S3A1", B = "S3B1")
  got <- wb_to_df(wb3, named_region = "MyRange", sheet = 4, colNames = FALSE)
  expect_equal(exp, got, ignore_attr = TRUE)

})


test_that("select_active_sheet", {

  wb <- wb_load(file = testfile_path("loadExample.xlsx"))

  exp <- structure(
    list(tabSelected = c("", "1", "", ""),
         workbookViewId = c("0", "0", "0", "0"),
         names = c("IrisSample", "testing", "mtcars", "mtCars Pivot")),
    row.names = c(NA, 4L), class = "data.frame")

  # testing is the selected sheet
  expect_identical(exp, wb_get_selected(wb))

  # change the selected sheet to IrisSample
  exp <- structure(
     list(tabSelected = c("1", "0", "0", "0"),
          workbookViewId = c("0", "0", "0", "0"),
          names = c("IrisSample", "testing", "mtcars", "mtCars Pivot")),
     row.names = c(NA, 4L), class = "data.frame")

  wb <- wb_set_selected(wb, "IrisSample")
  expect_identical(exp, wb_get_selected(wb))

  # get the active sheet
  expect_identical(2, wb_get_active_sheet(wb))

  # change the selected sheet to IrisSample
  wb <- wb_set_active_sheet(wb, sheet = "IrisSample")
  expect_identical(1, wb_get_active_sheet(wb))

})


test_that("dims_to_dataframe", {

  exp <- structure(
    list(A = c("A1", "A2"), C = c("C1", "C2")),
    row.names = 1:2,
    class = "data.frame"
  )
  got <- dims_to_dataframe("A1:A2;C1:C2", fill = TRUE)
  expect_equal(exp, got)

  got <- dims_to_dataframe("A1;A2;C1;C2", fill = TRUE)
  expect_equal(exp, got)

  exp <- list(c("A", "B"), "1")
  got <- dims_to_rowcol("A1;B1")
  expect_equal(exp, got)

})

test_that("dataframe_to_dims", {

  # dims_to_dataframe will always create a square
  df <- dims_to_dataframe("A1:D5;F1:F6;D8", fill = TRUE)
  dims <- dataframe_to_dims(df)
  df2 <- dims_to_dataframe(dims, fill = TRUE)
  expect_equal(df, df2)

})

test_that("handle 29Feb1900", {

  dates <- c("1900-02-28", "1900-03-01")
  as_date <- as.Date(dates)
  as_posix <- as.POSIXct(dates)

  exp <- c(59, 61)
  got <- conv_to_excel_date(as_date)
  expect_equal(exp, got)

  got <- conv_to_excel_date(as_posix)
  expect_equal(exp, got)

  expect_warning(
    conv_to_excel_date("x"),
    "could not convert x to Excel date. x is of class: character"
  )

  wb <- wb_workbook()$
    add_worksheet()$add_data(x = as_date)$
    add_worksheet()$add_data(x = as_posix)

  got <- wb_to_df(wb, sheet = 1, colNames = FALSE)$A
  expect_equal(as_date, got)

  got <- wb_to_df(wb, sheet = 2, colNames = FALSE)$A
  expect_equal(as_posix, got)

})


test_that("fillMergedCells works with dims", {

  # create data frame with emtpy second row
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = t(letters[1:4]), colNames = FALSE)$
    add_data(1, t(matrix(c(1:3, NA_real_), 4, 4)),
             startRow = 3, startCol = 1, colNames = FALSE)

  # merge rows 1 and 2 in each column
  wb$merge_cells(1, rows = 1:2, cols = 1)
  wb$merge_cells(1, rows = 1:2, cols = 2)
  wb$merge_cells(1, rows = 1:2, cols = 3)
  wb$merge_cells(1, rows = 1:2, cols = 4)

  # read from second column and fill merged cells
  got <- wb_to_df(wb, dims = "A2:D4", fillMergedCells = TRUE)

  exp <- c("a", "b", "c", "d")
  got <- names(got)
  expect_equal(exp, got)

})

test_that("improve date detection", {

  df <- wb_workbook() %>%
    wb_add_worksheet("Rawdata") %>%
    wb_add_data(x = Sys.Date(), colNames = FALSE) %>%
    wb_add_numfmt(numfmt = "[$-1070000]d/mm/yyyy;@") %>%
    wb_to_df(colNames = FALSE)

  exp <- Sys.Date()
  got <- df$A
  expect_equal(exp, got)

})

test_that("skip hidden columns and rows works", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)$
    set_col_widths(cols = c(1, 4, 6, 7, 9), hidden = TRUE)$
    set_row_heights(rows = c(3, 5, 8:30), hidden = TRUE)$
    add_data(dims = "M1", x = iris)

  dat <- wb_to_df(wb, dims = "A1:K33", skipHiddenRows = TRUE, skipHiddenCols = TRUE)

  exp <- c("2", "4", "6", "7", "31", "32", "33")
  got <- rownames(dat)
  expect_equal(exp, got)

  exp <- c("cyl", "disp", "drat", "vs", "gear", "carb")
  got <- names(dat)
  expect_equal(exp, got)

})

test_that("cols return order is correct", {

  wb <- wb_workbook()$add_worksheet()$add_data(dims = "B2", x = head(iris))

  exp <- structure(
    list(
      c(NA_real_, NA_real_, NA_real_, NA_real_, NA_real_, NA_real_),
      c(5.1, 4.9, 4.7, 4.6, 5, 5.4),
      c(3.5, 3, 3.2, 3.1, 3.6, 3.9)
    ),
    names = c(NA, "Sepal.Length", "Sepal.Width"),
    row.names = 3:8,
    class = "data.frame"
  )
  got <- wb_to_df(wb, cols = c(1:3))
  expect_equal(exp, got)

})
