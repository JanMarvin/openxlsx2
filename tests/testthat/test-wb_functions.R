test_that("wb_to_df", {

  ###########################################################################
  # numerics, dates, missings, bool and string
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  expect_silent(wb1 <- loadWorkbook(xlsxFile))

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
    row.names = c("2", "3", "4", "5", "6", "7", "8", "9", "10", "11"),
    class = "data.frame"
  )
  got <- wb_to_df(wb1)
  expect_equal(exp, got, check.attributes = FALSE)

  # do not convert first row to colNames
  got <- wb_to_df(wb1, colNames = FALSE)
  expect_equal(int2col(seq_along(got)), names(got))

  # do not try to identify dates in the data
  got <- wb_to_df(wb1, detectDates = FALSE)
  expect_equal(convertToExcelDate(df = exp["Var5"], date1904 = FALSE),
               got["Var5"])

  # return the underlying Excel formula instead of their values
  got <- wb_to_df(wb1, showFormula = TRUE)
  expect_equal("1/0", got$Var7[1])

  # read dimension withot colNames
  got <- wb_to_df(wb1, dims = "A2:C5", colNames = FALSE)
  test <- data.frame(A = c(TRUE, TRUE, TRUE, FALSE),
                     B = c(1, NA, 2, 2),
                     C = rep(NA_real_, 4))
  expect_equal(test, got, check.attributes = FALSE)

  # read selected cols
  got <- wb_to_df(wb1, cols = c(1:2, 7))
  expect_equal(exp[c(1,2,7)], got, check.attributes = FALSE)

  # read selected rows
  got <- wb_to_df(wb1, rows = c(1, 4, 6))
  got[c(4,7)] <- lapply(got[c(4,7)], as.character)
  expect_equal(exp[c(3,5),], got, check.attributes = FALSE)

  # convert characters to numerics and date (logical too?)
  got <- wb_to_df(wb1, convert = FALSE)
  chrs <- exp
  chrs[seq_along(chrs)] <- lapply(chrs[seq_along(chrs)], as.character)
  expect_equal(chrs, got, check.attributes = FALSE)

  # # erase empty Rows from dataset
  # not gonna test this :) just want to mention how blazing fast it is now.
  # got <- wb_to_df(wb1, sheet = 3, skipEmptyRows = TRUE)

  # erase rmpty Cols from dataset
  got <- wb_to_df(wb1, skipEmptyCols = TRUE)
  expect_equal(exp[c(1, 2, 4, 5, 6, 7, 8)], got, check.attributes = FALSE)

  # # convert first row to rownames
  # wb_to_df(wb1, sheet = 3, dims = "C6:G9", rowNames = TRUE)

  # define type of the data.frame
  got <- wb_to_df(wb1, cols = c(1, 4), types = c("Var1" = 0, "Var3" = 1))
  test <- exp[c("Var1", "Var3")]
  test["Var1"] <- lapply(test["Var1"], as.character)
  suppressWarnings(test["Var3"] <- lapply(test["Var3"], as.numeric))
  expect_equal(test, got, check.attributes = FALSE)

  # start in row 5
  got <- wb_to_df(wb1, startRow = 5, colNames = FALSE)
  test <- exp[4:10,]
  names(test) <- int2col(seq_along(test))
  test[c("D", "G", "H")] <- lapply(test[c("D", "G", "H")], as.numeric)
  expect_equal(test, got, check.attributes = FALSE)

  # na string
  got <- wb_to_df(wb1, na.strings = "")
  expect_equal("#N/A", got$Var7[2], check.attributes = FALSE)


  ###########################################################################
  # inlinestr
  xlsxFile <- system.file("extdata", "inline_str.xlsx", package = "openxlsx2")
  expect_silent(wb2 <- loadWorkbook(xlsxFile))

  exp <- data.frame(
    PairIndex = c(rep(1, 8), 2, 2), Drug1 = "abc", Drug2 = "def", Conc1 = 10000,
    Conc2 = c(10000, 3000, 1000, 300, 100, 30, 10, 0, 0, 10),
    Response = c(
      -1.79607109448082, 1.01028999064546, 0.449017773620206, 0,
      0.898035547240412, 0.112254443405051, 3.14312441534144,
      1.45930776426567, 1.68381665107577, -0.78578110383536 ),
    concUnit = "nM",
    stringsAsFactors = FALSE)
  rownames(exp) <- seq(2, nrow(exp)+ 1)
  # read dataset with inlinestr
  got <- wb_to_df(wb2)
  expect_equal(exp, got, check.attributes = FALSE)


  ###########################################################################
  # definedName // namedRegion
  xlsxFile <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
  expect_silent(wb3 <- loadWorkbook(xlsxFile))

  # read dataset with definedName (returns global first)
  exp <- data.frame(A = "S2A1", B = "S2B1")
  got <- wb_to_df(wb3, definedName = "MyRange", colNames = FALSE)
  expect_equal(exp, got, check.attributes = FALSE)

  # read definedName from sheet
  exp <- data.frame(A = "S3A1", B = "S3B1")
  got <- wb_to_df(wb3, definedName = "MyRange", sheet = 4, colNames = FALSE)
  expect_equal(exp, got, check.attributes = FALSE)

})
