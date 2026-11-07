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
    "names" = c("Var1", "Var2", NA, "Var3", "Var4", "Var5", "Var6", "Var7"),
    row.names = c(2:11),
    class = "data.frame"
  )
  got <- wb_to_df(wb1)
  expect_equal(got, exp, ignore_attr = TRUE)

  # do not convert first row to colNames
  got <- wb_to_df(wb1, col_names = FALSE)
  expect_equal(int2col(seq_along(got)), names(got))

  # do not try to identify dates in the data
  got <- wb_to_df(wb1, detect_dates = FALSE)
  expect_equal(convert_to_excel_date(df = exp["Var5"], date1904 = FALSE),
               got["Var5"])

  expect_error(convertToExcelDate(df = exp["Var5"], date1904 = FALSE), "defunct")

  # return the underlying Excel formula instead of their values
  got <- wb_to_df(wb1, show_formula = TRUE)
  expect_equal(got$Var7[1], "1/0")

  # read dimension without colNames
  got <- wb_to_df(wb1, dims = "A2:C5", col_names = FALSE)
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
  # got <- wb_to_df(wb1, sheet = 3, skip_empty_rows = TRUE)

  # erase rmpty Cols from dataset
  got <- wb_to_df(wb1, skip_empty_cols = TRUE)
  expect_equal(exp[c(1, 2, 4, 5, 6, 7, 8)], got, ignore_attr = TRUE)

  # # convert first row to rownames
  # wb_to_df(wb1, sheet = 3, dims = "C6:G9", row_names = TRUE)

  # define type of the data.frame
  got <- wb_to_df(wb1, cols = c(1, 4), types = c("Var1" = 0, "Var3" = 1))
  test <- exp[c("Var1", "Var3")]
  test["Var1"] <- lapply(test["Var1"], as.character)
  suppressWarnings(test["Var3"] <- lapply(test["Var3"], function(x) as.numeric(replace(x, x == "#NUM!", "NaN"))))
  expect_equal(test, got, ignore_attr = TRUE)

  # start in row 5
  got <- wb_to_df(wb1, start_row = 5, col_names = FALSE)
  test <- exp[4:10, ]
  names(test) <- int2col(seq_along(test))
  test[c("D", "G", "H")] <- lapply(test[c("D", "G", "H")], as.numeric)
  expect_equal(got, test, ignore_attr = TRUE)

  # na string
  got <- wb_to_df(wb1, na.strings = "")
  expect_equal(got$Var7[2], "#N/A", ignore_attr = TRUE)


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
  expect_equal(got, exp, ignore_attr = TRUE)


  ###########################################################################
  # named_region // namedRegion
  xlsxFile <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
  expect_silent(wb3 <- wb_load(xlsxFile))

  # read dataset with named_region (returns global first)
  got <- wb_to_df(wb3, named_region = "MyRange", col_names = FALSE)
  exp <- data.frame(A = "S2A1", B = "S2B1", stringsAsFactors = FALSE)
  expect_equal(got, exp, ignore_attr = TRUE)

  # read named_region from sheet
  got <- wb_to_df(wb3, named_region = "MyRange", sheet = 4, col_names = FALSE)
  exp <- data.frame(A = "S3A1", B = "S3B1", stringsAsFactors = FALSE)
  expect_equal(got, exp, ignore_attr = TRUE)

  ## read blank cells and avoid warnings and errors
  wb <- wb_workbook()$add_worksheet()$add_data(x = 1, col_names = FALSE)

  exp <- structure(list(character(0), character(0)),
                  names = c(NA_character_, NA_character_),
                  row.names = integer(0), class = "data.frame")
  got <- wb$to_df(dims = "D2:E5", skip_empty_rows = TRUE)
  expect_equal(got, exp)

  exp <- structure(list(), names = character(0), row.names = 3:5, class = "data.frame")
  got <- wb$to_df(dims = "D2:E5", skip_empty_cols = TRUE)
  expect_equal(got, exp)

})


test_that("select_active_sheet", {

  wb <- wb_load(file = testfile_path("loadExample.xlsx"))

  exp <- structure(
    list(tabSelected = c("", "1", "", ""),
         workbookViewId = c("0", "0", "0", "0"),
         names = c("IrisSample", "testing", "mtcars", "mtCars Pivot")),
    row.names = c(NA, 4L), class = "data.frame")

  # testing is the selected sheet
  expect_identical(wb_get_selected(wb), exp)

  # change the selected sheet to IrisSample
  exp <- structure(
     list(tabSelected = c("1", "0", "0", "0"),
          workbookViewId = c("0", "0", "0", "0"),
          names = c("IrisSample", "testing", "mtcars", "mtCars Pivot")),
     row.names = c(NA, 4L), class = "data.frame")

  wb <- wb_set_selected(wb, "IrisSample")
  expect_identical(wb_get_selected(wb), exp)

  # get the active sheet
  expect_identical(wb_get_active_sheet(wb), 2)

  # change the selected sheet to IrisSample
  wb <- wb_set_active_sheet(wb, sheet = "IrisSample")
  expect_identical(wb_get_active_sheet(wb), 1)

})

test_that("get_selected respects set_order", {

  wb <- wb_workbook()
  for (s in c("Data", "Notes", "Disclaimer")) wb$add_worksheet(s)
  wb$set_order(c("Disclaimer", "Data", "Notes"))

  got <- wb$get_selected()
  expect_identical(got$names, c("Disclaimer", "Data", "Notes"))
  expect_identical(got$tabSelected, c("0", "1", "0"))

  wb$set_selected("Disclaimer")
  expect_identical(wb$get_selected()$tabSelected, c("1", "0", "0"))

})

test_that("selecting several sheets keeps the current sheet", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_selected(c("A", "B"))
  wb$set_order(c("C", "B", "A"))
  wb$add_data(x = "on C")
  expect_identical(wb_to_df(wb, sheet = "C", col_names = FALSE)[[1]], "on C")

})

test_that("set_active_sheet by name respects set_order", {

  wb <- wb_workbook()
  for (s in c("Data", "Notes", "Summary", "Disclaimer")) wb$add_worksheet(s)
  wb$set_order(c("Disclaimer", "Data", "Notes", "Summary"))

  wb$set_active_sheet()
  expect_identical(wb$get_active_sheet(), 1)

  wb$set_active_sheet("Disclaimer")
  expect_identical(wb$get_active_sheet(), 1)

  wb$set_active_sheet("Summary")
  expect_identical(wb$get_active_sheet(), 4)

  wb$set_active_sheet(1)
  expect_identical(wb$get_active_sheet(), 1)

  tmp <- temp_xlsx()
  wb$set_active_sheet("Notes")
  wb$save(tmp)
  wb2 <- wb_load(tmp)
  expect_identical(wb2$get_active_sheet(), 3)
  expect_identical(unname(wb2$get_sheet_names()[3]), "Notes")

})

test_that("local sheet ids follow set_order", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "A", "B"))

  wb$add_named_region(sheet = "C", dims = "A1", name = "nc", local_sheet = TRUE)
  wb$set_page_setup(sheet = "C", print_title_rows = 1)
  wb$add_data(sheet = "C", x = data.frame(x = 1), with_filter = TRUE)

  dn <- wb$get_named_regions(builtins = TRUE)
  expect_identical(unique(dn$localSheetId), "0")
  expect_identical(unique(dn$sheets), "C")

  wb$set_active_sheet("A")
  wb$set_order(c("B", "A", "C"))
  dn <- wb$get_named_regions(builtins = TRUE)
  expect_identical(unique(dn$localSheetId), "2")
  expect_identical(wb$get_active_sheet(), 2)

  wb$set_order(c("A", "C", "B"))
  dn <- wb$get_named_regions(builtins = TRUE)
  expect_identical(unique(dn$localSheetId), "1")
  expect_identical(wb$get_active_sheet(), 1)

})

test_that("remove_worksheet shifts written positions", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$add_named_region(sheet = "C", dims = "A1", name = "nc", local_sheet = TRUE)
  wb$remove_worksheet("A")
  expect_identical(wb$get_named_regions()$localSheetId, "1")

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "A", "B"))
  wb$set_active_sheet("A")
  wb$remove_worksheet("C")
  expect_identical(wb$get_active_sheet(), 1)

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_active_sheet("B")
  wb$remove_worksheet("C")
  expect_identical(wb$get_active_sheet(), 2)

  wb$set_active_sheet("B")
  wb$remove_worksheet("B")
  expect_identical(wb$get_active_sheet(), 1)

})

test_that("set_sheet_names respects set_order and keeps attributes", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "A", "B"))
  wb$add_named_region(sheet = "C", dims = "A1", name = "nc")
  wb$set_sheet_visibility("B", "hidden")

  wb$set_sheet_names("C", "Z")
  wb$set_sheet_names("B", "Q")
  expect_identical(unname(wb$get_sheet_names()), c("Z", "A", "Q"))
  expect_identical(wb$get_sheet_visibility(), c("visible", "visible", "hidden"))
  expect_identical(xml_value(wb$workbook$definedNames, "definedName"), "'Z'!$A$1:$A$1")

  wb <- wb_workbook()
  for (s in c("A", "B")) wb$add_worksheet(s)
  wb$add_named_region(sheet = "A", dims = "A1", name = "na")
  wb$add_named_region(sheet = "B", dims = "A1", name = "nb")
  wb$set_sheet_names(c("A", "B"), c("X", "Y"))
  expect_identical(
    xml_value(wb$workbook$definedNames, "definedName"),
    c("'X'!$A$1:$A$1", "'Y'!$A$1:$A$1")
  )

})

test_that("sheet visibility follows set_order", {

  wb <- wb_workbook()
  for (s in c("hidden", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "hidden", "B"))
  wb$set_sheet_visibility("C", "veryHidden")
  expect_identical(wb$get_sheet_visibility(), c("veryHidden", "visible", "visible"))

  tmp <- temp_xlsx()
  wb$save(tmp)
  wb2 <- wb_load(tmp)
  expect_identical(wb2$get_sheet_visibility(), c("veryHidden", "visible", "visible"))
  expect_identical(wb2$get_active_sheet(), 2)

  expect_warning(
    wb$set_sheet_visibility(c("hidden", "B"), c(FALSE, FALSE)),
    "1 visible worksheet"
  )
  expect_identical(wb$get_sheet_visibility(), c("visible", "hidden", "hidden"))

})

test_that("defined names are not overwritten or duplicated", {

  wb <- wb_workbook()
  for (s in c("A", "B")) wb$add_worksheet(s)
  wb$add_named_region(sheet = "A", dims = "B2", name = "aaa", local_sheet = TRUE)
  wb$add_data(sheet = "B", x = data.frame(v = 1:2), dims = "D1", with_filter = TRUE)
  wb$add_data(sheet = "B", x = data.frame(v = 1:2), dims = "F1", with_filter = TRUE)
  wb$set_page_setup(sheet = "B", print_title_rows = 1)
  wb$set_page_setup(sheet = "B", print_title_rows = 2)

  dn <- rbindlist(xml_attr(wb$workbook$definedNames, "definedName"))
  expect_identical(sort(dn$name), sort(c("aaa", "_xlnm._FilterDatabase", "_xlnm.Print_Titles")))
  expect_true(any(grepl("$F$1", wb$workbook$definedNames, fixed = TRUE)))
  expect_true(any(grepl("$2:$2", wb$workbook$definedNames, fixed = TRUE)))

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_page_setup(sheet = "A", print_title_rows = 1)
  wb$add_named_region(sheet = "C", dims = "B2", name = "keep", local_sheet = TRUE)
  wb$remove_worksheet("B")
  dn <- rbindlist(xml_attr(wb$workbook$definedNames, "definedName"))
  expect_identical(dn$name, c("_xlnm.Print_Titles", "keep"))
  expect_identical(dn$localSheetId, c("0", "1"))

})

test_that("selection by position follows set_order", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "A", "B"))
  wb$set_selected(1)
  sel <- wb$get_selected()
  expect_identical(sel$names[sel$tabSelected == "1"], "C")

  wb <- wb_workbook()
  wb$add_worksheet("A")
  expect_identical(wb$get_selected()$names, "A")

})

test_that("set_sheet_names without old follows set_order", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "A", "B"))
  wb$set_sheet_names(new = paste0(wb$get_sheet_names(), "x"))
  expect_identical(unname(wb$get_sheet_names()), c("Cx", "Ax", "Bx"))

  wb$set_sheet_names("Cx", "cx")
  expect_identical(unname(wb$get_sheet_names()), c("cx", "Ax", "Bx"))

})

test_that("removing sheets keeps tables, first sheet and app.xml in sync", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$add_data_table(sheet = "B", x = data.frame(t = 1:2), table_name = "tb")
  wb$remove_worksheet("A")
  expect_identical(wb$get_tables("B")$tab_name, "tb")
  expect_length(wb$get_tables("C"), 2)
  expect_equal(NROW(wb$get_tables("C")), 0)

  wb$remove_worksheet("B")
  expect_null(wb$get_named_regions(tables = TRUE))

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_bookview(first_sheet = 2, active_tab = 2)
  wb$remove_worksheet("C")
  expect_identical(wb$get_bookview()$firstSheet, "1")

  wb <- wb_workbook()
  wb$add_worksheet("A")
  wb$add_chartsheet("C")
  wb$add_worksheet("B")
  wb$set_order(c("C", "B", "A"))
  tmp <- temp_xlsx()
  wb$save(tmp)
  dir <- tempfile()
  unzip(tmp, files = "docProps/app.xml", exdir = dir)
  app <- read_xml(file.path(dir, "docProps/app.xml"))
  parts <- xml_node(app, "Properties", "TitlesOfParts")
  expect_true(grepl("<vt:lpstr>B</vt:lpstr><vt:lpstr>A</vt:lpstr><vt:lpstr>C</vt:lpstr>", parts, fixed = TRUE))

  wb <- wb_workbook()
  wb$add_chartsheet("C")
  wb$save(tmp)
  expect_identical(unname(wb_load(tmp)$get_sheet_names()), "C")

})

test_that("clean_sheet removes tables on sheets with xml entities", {

  wb <- wb_workbook()
  wb$add_worksheet("a&b")
  wb$add_data_table(x = data.frame(t = 1:2), table_name = "tb")
  wb$clean_sheet(sheet = "a&b", tables = TRUE)
  expect_equal(NROW(wb$get_tables("a&b")), 0)

})

test_that("clone and copy_cells follow set_order", {

  wb <- wb_workbook()
  for (s in c("A", "B")) wb$add_worksheet(s)
  wb$add_data(sheet = "B", x = "on B", dims = "A1", col_names = FALSE)
  wb$add_named_region(sheet = "B", dims = "B2", name = "loc", local_sheet = TRUE)
  wb$set_order(c("B", "A"))

  expect_warning(wb$clone_worksheet("B", "C"), "named regions")
  dn <- rbindlist(xml_attr(wb$workbook$definedNames, "definedName"))
  expect_identical(dn$localSheetId[dn$name == "loc"], c("0", "2"))

  wb <- wb_workbook()
  for (s in c("A", "B")) wb$add_worksheet(s)
  wb$set_order(c("B", "A"))
  wb$workbook$definedNames <- '<definedName localSheetId="0" name="k">1</definedName>'
  expect_warning(wb$clone_worksheet("B", "C"), "named regions")
  expect_identical(rbindlist(xml_attr(wb$workbook$definedNames, "definedName"))$localSheetId, c("0", "2"))

  wb <- wb_workbook()
  for (s in c("A", "B")) wb$add_worksheet(s)
  wb$add_data(sheet = "B", x = "on B", dims = "A1", col_names = FALSE)
  wb$set_order(c("B", "A"))
  dat <- wb_data(wb, sheet = "B", dims = "A1", col_names = FALSE)
  expect_identical(attr(dat, "sheet"), "B")
  wb$copy_cells(sheet = "A", dims = "C3", data = dat, as_ref = TRUE)
  expect_identical(wb$worksheets[[1]]$sheet_data$cc$f[wb$worksheets[[1]]$sheet_data$cc$r == "C3"], "'B'!A1")

})

test_that("print titles with columns follow set_order", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "A", "B"))
  wb$set_page_setup(sheet = "C", print_title_cols = 1:2)
  wb$set_page_setup(sheet = "A", print_title_rows = 1, print_title_cols = 1)

  dn <- rbindlist(xml_attr(wb$workbook$definedNames, "definedName"))
  val <- xml_value(wb$workbook$definedNames, "definedName")
  expect_identical(dn$localSheetId, c("0", "1"))
  expect_identical(val, c("'C'!$A:$B", "'A'!$A:$A,'A'!$1:$1"))

})

test_that("named regions survive deleted tables and broken references", {

  wb <- wb_workbook()
  wb$add_worksheet("A")
  wb$add_named_region(dims = "B2", name = "nr")
  wb$add_data_table(x = data.frame(t = 1:2), dims = "D1", table_name = "tb")
  wb$remove_tables(table = "tb")
  expect_identical(wb$get_named_regions(tables = TRUE)$name, "nr")

  wb$workbook$definedNames <- c(wb$workbook$definedNames, "<definedName name=\"gone\">#REF!$A$1</definedName>")
  wb$remove_named_region(sheet = "A")
  expect_identical(wb$workbook$definedNames, "<definedName name=\"gone\">#REF!$A$1</definedName>")

})

test_that("hidden sheets are unselected after set_order", {

  wb <- wb_workbook()
  for (s in c("A", "B", "C")) wb$add_worksheet(s)
  wb$set_order(c("C", "A", "B"))
  wb$set_selected(c("C", "B"))
  wb$set_sheet_visibility("B", "hidden")

  tmp <- temp_xlsx()
  wb$save(tmp)
  sel <- wb_load(tmp)$get_selected()
  expect_identical(sel$names[sel$tabSelected == "1"], "C")

})

test_that("sheet names with xml entities work with named regions", {

  for (nm in c("a&b", "x<y", "q\"t", "&gt;", "&amp;", "&lt;&gt;")) {

    wb <- wb_workbook()
    wb$add_worksheet(nm)
    wb$add_worksheet("B")
    wb$add_data(sheet = nm, x = data.frame(v = 1:2), dims = "B2", name = "rng")
    wb$add_named_region(sheet = nm, dims = "D4", name = "loc", local_sheet = TRUE)

    expect_identical(wb_get_named_regions(wb)$sheet, c(1L, 1L))

    wb$add_worksheet("C")
    wb$remove_worksheet("B")
    expect_length(wb$workbook$definedNames, 2)
    expect_identical(wb_to_df(wb, named_region = "rng")$v, c(1, 2))

    wb$set_sheet_names("C", paste0(nm, "x"))
    expect_identical(unname(wb$get_sheet_names()), c(nm, paste0(nm, "x")))
    wb$add_named_region(sheet = paste0(nm, "x"), dims = "B2", name = "other")

    wb$remove_named_region(sheet = nm)
    expect_identical(wb_get_named_regions(wb)$name, "other")

    tmp <- temp_xlsx()
    wb$save(tmp)
    wb2 <- wb_load(tmp)
    expect_identical(unname(wb2$get_sheet_names()), c(nm, paste0(nm, "x")))
  }

})


test_that("dims_to_dataframe", {

  exp <- structure(
    list(A = c("A1", "A2"), C = c("C1", "C2")),
    row.names = 1:2,
    class = "data.frame"
  )
  got <- dims_to_dataframe("A1:A2;C1:C2", fill = TRUE)
  expect_equal(exp, got[c("A", "C")])

  got <- dims_to_dataframe("A1;A2,C1;C2", fill = TRUE)
  expect_equal(exp, got[c("A", "C")])

  got <- dims_to_dataframe("A1:A2;C1:C2", fill = TRUE, empty_rm = TRUE)
  expect_equal(exp, got)

  got <- dims_to_dataframe("A1;A2,C1;C2", fill = TRUE, empty_rm = TRUE)
  expect_equal(exp, got)

  exp <- list(col = c("A", "B"), row = "1")
  got <- dims_to_rowcol("A1;B1")
  expect_equal(exp, got)

})

test_that("dataframe_to_dims", {

  # dims_to_dataframe will always create a square
  df <- dims_to_dataframe("A1:D5;F1:F6;D8", fill = TRUE)
  dims <- dataframe_to_dims(df, dim_break = TRUE)
  df2 <- dims_to_dataframe(dims, fill = TRUE)
  expect_equal(df, df2)

})

test_that("handle 29Feb1900", {

  dates <- c("1900-02-28", "1900-03-01")
  as_date <- as.Date(dates)
  as_posix <- as.POSIXct(dates, tz = "UTC")

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

  got <- wb_to_df(wb, sheet = 1, col_names = FALSE)$A
  expect_equal(as_date, got)

  got <- wb_to_df(wb, sheet = 2, col_names = FALSE)$A
  expect_equal(as_posix, got)

})


test_that("fillMergedCells works with dims", {

  # create data frame with empty second row
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = t(letters[1:4]), col_names = FALSE)$
    add_data(1, t(matrix(c(1:3, NA_real_), 4, 4)),
             start_row = 3, start_col = 1, col_names = FALSE)

  # merge rows 1 and 2 in each column
  wb$merge_cells(1, wb_dims(rows = 1:2, cols = 1))
  wb$merge_cells(1, wb_dims(rows = 1:2, cols = 2))
  wb$merge_cells(1, wb_dims(rows = 1:2, cols = 3))
  wb$merge_cells(1, wb_dims(rows = 1:2, cols = 4))

  # read from second column and fill merged cells
  got <- wb_to_df(wb, dims = "A2:D4", fill_merged_cells = TRUE)

  exp <- c("a", "b", "c", "d")
  expect_equal(names(got), exp)

})

test_that("improve date detection", {

  df <- wb_workbook()$
    add_worksheet("Rawdata")$
    add_data(x = Sys.Date(), col_names = FALSE)$
    add_numfmt(numfmt = "[$-1070000]d/mm/yyyy;@")$
    to_df(col_names = FALSE)

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

  dat <- wb_to_df(wb, dims = "A1:K33", skip_hidden_rows = TRUE, skip_hidden_cols = TRUE)

  exp <- c("2", "4", "6", "7", "31", "32", "33")
  got <- rownames(dat)
  expect_equal(rownames(dat), exp)

  exp <- c("cyl", "disp", "drat", "vs", "gear", "carb")
  expect_equal(names(dat), exp)

})

test_that("test that skip_hidden_cols works with data_only", {
  tmp <- temp_xlsx()
  on.exit(unlink(tmp), add = TRUE)

  mm <- matrix(1, 3, 3)
  mm[2, ] <- 2
  mm[, 2] <- 2

  wb_workbook()$
    add_worksheet()$
    add_data(x = mm)$
    set_row_heights(rows = 3, hidden = TRUE)$
    set_col_widths(cols = "B", hidden = TRUE)$
    save(tmp)

  wb <- wb_load(tmp, data_only = TRUE)

  df1 <- wb_to_df(wb, skip_hidden_rows = TRUE)
  df2 <- wb_to_df(wb, skip_hidden_cols = TRUE)
  df3 <- wb_to_df(wb, skip_hidden_rows = TRUE, skip_hidden_cols = TRUE)

  df <- data.frame(mm)
  names(df) <- paste0("V", seq_along(df))
  row.names(df) <- as.integer(2:4)

  sel <- c(1, 3)
  edf1 <- df[sel, ]
  edf2 <- df[, sel]
  edf3 <- df[sel, sel]

  expect_equal(edf1, df1)
  expect_equal(edf2, df2)
  expect_equal(edf3, df3)
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

test_that("missings cells are returned", {

  wb <- wb_workbook()$add_worksheet()$add_data(x = "a", dims = "Z100")

  expect_silent(got <- wb_to_df(wb, col_names = FALSE, rows = 1, cols = 1))
  expect_silent(got <- wb_to_df(wb, col_names = FALSE, rows = c(1, 4, 10), cols = c(2, 5, 10)))

  expect_silent(got <- wb_to_df(wb, col_names = FALSE, rows = c(101), cols = c(27)))

})

test_that("dims with separator work", {

  wb <- wb_workbook()$
    # this picks the top left corner of dims to write the data frame
    add_worksheet()$add_data(dims = "I2:J2;A1:B2;G5:H6", x = matrix(1:8, 4, 2))$
    add_worksheet()$add_data(dims = "I2:J2;A1:B2;G5:H6", x = matrix(1:8, 4, 2), enforce = TRUE)

  exp <- c("A1", "B1", "A2", "B2", "A3", "B3", "A4", "B4", "A5", "B5")
  got <- wb$worksheets[[1]]$sheet_data$cc$r
  expect_equal(exp, got)

  exp <- c("I2", "J2", "A1", "B1", "A2", "B2", "G5", "H5", "G6", "H6")
  got <- wb$worksheets[[2]]$sheet_data$cc$r
  expect_equal(exp, got)

  # write a workbook with a few colored cells
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(dims = wb_dims(x = mtcars), x = mtcars, enforce = TRUE)$
    add_worksheet()$
    add_data(dims = "A1:K20,A22:K34", x = mtcars, enforce = TRUE)$
    add_worksheet()$
    add_data(dims = "I2:J2;A1:B2;G5:H6", x = matrix(1:8, 4, 2), enforce = TRUE)

  # sheet 1
  df <- wb_to_df(wb, sheet = 1)
  expect_equal(df, mtcars, ignore_attr = TRUE)

  # sheet 2
  df <- wb_to_df(wb, sheet = 2)
  expect_true(all(is.na(df[20, ])))
  expect_equal(df[-20, ], mtcars, ignore_attr = TRUE)

  # sheet 3
  df <- wb_to_df(wb, sheet = 3, col_names = FALSE)
  ll <- list(
    df["2", c("I", "J")],
    df[c("1", "2"), c("A", "B")],
    df[c("5", "6"), c("G", "H")]
  )

  for (i in seq_along(ll))
    names(ll[[i]]) <- c("V1", "V2")

  exp <- data.frame(
    V1 = c("V1", 1, 2, 3, 4),
    V2 = c("V2", 5, 6, 7, 8),
    stringsAsFactors = FALSE
  )

  got <- do.call("rbind", ll)
  expect_equal(exp, got, ignore_attr = TRUE)

  wb$add_worksheet()$add_formula(dims = "D1;F2", x = c("SUM(A1:C1)", "SUM(A1)"), enforce = TRUE)
  exp <- c("D1", "F2")
  got <- wb$worksheets[[4]]$sheet_data$cc$r
  expect_equal(exp, got)

  # write to different dims
  dat <- as.data.frame(matrix(1:15, ncol = 3))

  wb <- wb_workbook()

  # write columns separately
  wb$add_worksheet()$
    add_data(dims = "A1:A6", x = dat[, 1, drop = FALSE],
            enforce = TRUE, col_names = TRUE)$
    add_data(dims = "C2:C4,C6:C7", x = dat[, 2, drop = FALSE],
            enforce = TRUE, col_names = FALSE)

  # write columns as unique cells
  wb$add_worksheet()$
    add_data(dims = "A1,B2,C3,D4,E5,F6,C1,D2,E3,F4,G5,H6,I1:I6", x = dat,
            enforce = TRUE, col_names = TRUE)

  # write columns as two column ranges
  wb$add_worksheet()$
    add_data(dims = "A1:A6,C1:C6,E1:E6", x = dat,
            enforce = TRUE, col_names = TRUE)

  # write columns as two row ranges
  wb$add_worksheet()$
    add_data(dims = "A1:C3,B5:D7", x = dat,
            enforce = TRUE, col_names = TRUE)


  exp <- c("A1", "A2", "C2", "A3", "C3", "A4", "C4", "A5", "C5", "A6", "C6", "C7")
  got <- wb$worksheets[[1]]$sheet_data$cc$r
  expect_equal(exp, got)

  exp <- c("A1", "C1", "I1", "B2", "D2", "I2", "C3", "E3", "I3", "D4",
          "F4", "I4", "E5", "G5", "I5", "F6", "H6", "I6")
  got <- wb$worksheets[[2]]$sheet_data$cc$r
  expect_equal(exp, got)

  exp <- c("A1", "C1", "E1", "A2", "C2", "E2", "A3", "C3", "E3", "A4",
          "C4", "E4", "A5", "C5", "E5", "A6", "C6", "E6")
  got <- wb$worksheets[[3]]$sheet_data$cc$r
  expect_equal(exp, got)

  exp <- c("A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3", "B5",
          "C5", "D5", "B6", "C6", "D6", "B7", "C7", "D7")
  got <- wb$worksheets[[4]]$sheet_data$cc$r
  expect_equal(exp, got)

})

test_that("improve non consecutive dims", {

  # dims <- "B7:B9,C6:C10,D5:D11,E5:E12,F6:F13,G7:G14,H6:H13,I5:I12,J5:J11,K6:K10,L7:L9"
  dims <- "B7:B9,C6:C10,D5:D11,E5:E12,G7:G14,H6:H13,I5:I12,J5:J11,K6:K10,L7:L9"

  wb1 <- wb_workbook()$
    add_worksheet(grid_lines = FALSE)$
    set_col_widths(cols = "A:M", widths = 2)$
    add_fill(dims = dims, color = wb_color("red"))

  wb2 <- wb_workbook()$
    add_worksheet(grid_lines = FALSE)$
    set_col_widths(cols = "A:M", widths = 2)$
    add_data(x = "", dims = "B5:L14")$
    add_fill(dims = dims, color = wb_color("red"))

  expect_true(wb1$worksheets[[1]]$dimension == wb2$worksheets[[1]]$dimension)

  # TODO might want to exclude the empty cells here
  exp <- dims_to_dataframe(dims, fill = TRUE)
  exp <- unname(unlist(exp[exp != ""]))
  got <- wb1$worksheets[[1]]$sheet_data$cc$r[wb1$worksheets[[1]]$sheet_data$cc$c_s != ""]
  expect_identical(sort(got), sort(exp))

  got <- wb2$worksheets[[1]]$sheet_data$cc$r[wb2$worksheets[[1]]$sheet_data$cc$c_s != ""]
  expect_identical(sort(got), sort(exp))

  ### Test rowwise
  # dims <- "D5:E5,I5:J5,C6:F6,H6:K6,B7:L9,C10:K10,D11:J11,E12:I12,F13:H13,G14"
  dims <- "D5:E5,I5:J5,B7:L9,C10:K10,E12:I12,F13:H13,G14"

  wb3 <- wb_workbook()$
    add_worksheet(grid_lines = FALSE)$
    set_col_widths(cols = "A:M", widths = 2)$
    add_fill(dims = dims, color = wb_color("red"))

  wb4 <- wb_workbook()$
    add_worksheet(grid_lines = FALSE)$
    set_col_widths(cols = "A:M", widths = 2)$
    add_data(x = "", dims = "B5:L14")$
    add_fill(dims = dims, color = wb_color("red"))

  expect_true(wb3$worksheets[[1]]$dimension == wb4$worksheets[[1]]$dimension)

  exp <- dims_to_dataframe(dims, fill = TRUE)
  exp <- unname(unlist(exp[exp != ""]))
  got <- wb3$worksheets[[1]]$sheet_data$cc$r[wb3$worksheets[[1]]$sheet_data$cc$c_s != ""]
  expect_identical(sort(got), sort(exp))

  got <- wb4$worksheets[[1]]$sheet_data$cc$r[wb4$worksheets[[1]]$sheet_data$cc$c_s != ""]
  expect_identical(sort(got), sort(exp))
})

test_that("reading equal sized ranges works", {

  df <- as.data.frame(matrix(rep(1:4, 4), 4, 4, byrow = TRUE))

  wb <- wb_workbook()$add_worksheet()$add_data(x = df)

  exp <- df[, -2]
  got <- wb_to_df(wb, dims = "A1:A5,C1:D5")
  expect_equal(exp, got, ignore_attr = TRUE)

  got <- wb_to_df(wb, dims = "A1,C1:D1,A2:A5,C2:D5")
  expect_equal(exp, got, ignore_attr = TRUE)

})

test_that("creating a formula matrix works", {

  df <- matrix(
    1:100, ncol = 10, nrow = 10
  )

  wb <-  wb_workbook()$add_worksheet()$add_data(x = df)

  wb$add_formula(
    x      = "=$A2/B$2",
    dims   = wb_dims(x = df, from_row = 13, col_names = FALSE),
    shared = TRUE
  )

  exp <- c(210, 10)
  got <- dim(wb$worksheets[[1]]$sheet_data$cc)
  expect_equal(exp, got)

})

test_that("writing formula dataframes works", {

  df <- matrix(
    1:100, ncol = 10, nrow = 10
  )

  fml_df <- dims_to_dataframe(wb_dims(x = df, col_names = FALSE, from_row = 2), fill = TRUE)

  wb <-  wb_workbook()$add_worksheet()$add_data(x = df)

  wb$add_formula(
    x      = fml_df,
    dims   = wb_dims(x = df, from_row = 13, col_names = FALSE)
  )

  exp <- c(210, 10)
  got <- dim(wb$worksheets[[1]]$sheet_data$cc)
  expect_equal(exp, got)

})

test_that("dims order is respected", {
  wb <- write_xlsx(x = mtcars)
  exp <- wb$to_df(dims = "C1:A4,D1:E4")
  got <- mtcars[1:3, c(3:1, 4:5)]
  expect_equal(exp, got, ignore_attr = TRUE)

  rc <- dims_to_rowcol("C3:A1")
  expect_equal(rc$col, c("C", "B", "A"))
  expect_equal(rc$row, c("3", "2", "1"))
})

test_that("convert works", {

  df <- data.frame(
    is_active  = c(TRUE, FALSE, TRUE, NA, FALSE),
    count      = c(10L, 25L, NA, 7L, 15L),
    measure    = c(1.23, NA, 4.56, 7.89, 0.01),
    event_date = as.Date(c("2023-01-01", "2023-02-15", NA, "2023-05-20", "2023-12-31")),
    timestamp  = as.POSIXct(c("2023-01-01 10:00:00", NA, "2023-03-10 14:30:00",
                              "2023-06-01 08:15:00", "2023-12-25 12:00:00"), tz = "UTC"),
    daily_time = structure(c(30600, 35100, NA, 40500, 43200), units = "secs", class = c("hms", "difftime")),
    category   = c("A", "B", "C", NA, "E"),
    stringsAsFactors = FALSE
  )

  wb <- write_xlsx(x = df)

  df_u <- wb_to_df(wb, convert = FALSE)
  df_c <- wb_to_df(wb, convert = TRUE)

  expect_equal(dim(df_u), dim(df_c))
  expect_equal(is.na(df_u), is.na(df_c))

  df_u <- wb_to_df(wb, convert = FALSE, col_names = TRUE, rows = 2)
  df_c <- wb_to_df(wb, convert = TRUE, col_names = TRUE, rows = 2)

  expect_equal(dim(df_u), dim(df_c))

})

test_that("dims_to_dataframe() rejects NA dims", {

  expect_error(dims_to_dataframe(NA), "character vector of cell ranges")
  expect_error(dims_to_dataframe(NA_character_), "character vector of cell ranges")
  expect_error(dims_to_dataframe(1L), "character vector of cell ranges")

  wb <- wb_workbook()$add_worksheet()
  expect_error(wb$add_ignore_error(dims = NA, number_stored_as_text = TRUE), "character vector of cell ranges")

})
