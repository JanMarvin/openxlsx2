test_that("write_formula", {

  set.seed(123)
  df <- data.frame(C = rnorm(10), D = rnorm(10))

  # array formula for a single cell
  exp <- structure(
    list(r = "E2", row_r = "2", c_r = "E", c_s = "",
         c_t = "", c_cm = "",
         c_ph = "", c_vm = "",
         v = "", f = "SUM(C2:C11*D2:D11)",
         f_t = "array", f_ref = "E2",
         f_ca = "", f_si = "",
         is = "", typ = "11"),
    row.names = 23L, class = "data.frame")

  # write data add array formula later
  wb <- wb_workbook()
  wb <- wb_add_worksheet(wb, "df")
  wb$add_data("df", df, startCol = "C")
  write_formula(wb, "df", startCol = "E", startRow = 2,
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)

  cc <- wb$worksheets[[1]]$sheet_data$cc
  got <- cc[cc$row_r == "2" & cc$c_r == "E", ]
  expect_equal(exp[1:16], got[1:16])


  rownames(exp) <- 1L
  # write formula first add data later
  wb <- wb_workbook()
  wb <- wb_add_worksheet(wb, "df")
  write_formula(wb, "df", startCol = "E", startRow = 2,
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)
  wb$add_data("df", df, startCol = "C")

  cc <- wb$worksheets[[1]]$sheet_data$cc
  got <- cc[cc$row_r == "2" & cc$c_r == "E", ]
  expect_equal(exp[1:11], got[1:11])

})

test_that("silent with numfmt option", {

  wb <- wb_workbook()
  wb$add_worksheet("S1")
  wb$add_worksheet("S2")

  wb$add_data_table("S1", x = iris)
  wb$add_data_table("S2",
                    x = mtcars, dims = "B3", rowNames = TRUE,
                    tableStyle = "TableStyleLight9")

  # [1:4] to ignore factor
  expect_equal(iris[1:4], wb_to_df(wb, "S1")[1:4], ignore_attr = TRUE)
  expect_equal(iris[1:4], wb_to_df(wb, "S1")[1:4], ignore_attr = TRUE)

  # handle rownames
  got <- wb_to_df(wb, "S2", rowNames = TRUE)
  attr(got, "tt") <- NULL
  attr(got, "types") <- NULL
  expect_equal(mtcars, got)
  expect_equal(rownames(mtcars), rownames(got))

})

test_that("test options", {

  ops <- options()
  tmp <- temp_xlsx()
  wb_workbook()$add_worksheet("Sheet 1")$add_data("Sheet 1", cars)
  ops2 <- options()

  # adding data to the worksheet should not alter the global options
  expect_equal(ops, ops2)

})

test_that("missing x is caught early [246]", {
  expect_error(
    wb_workbook()$add_data(mtcars),
    "`x` is missing"
  )
})

test_that("update_cells", {

  ## exactly the same
  data <- mtcars
  wb <- wb_workbook()$add_worksheet()$add_data(x = data)
  cc1 <- wb$worksheets[[1]]$sheet_data$cc

  wb$add_data(x = data)
  cc2 <- wb$worksheets[[1]]$sheet_data$cc

  expect_equal(cc1, cc2)

  ## write na.strings
  data <- matrix(NA, 2, 2)
  wb <- wb_workbook()$add_worksheet()$add_data(x = data)$add_data(x = data, na.strings = "N/A")

  exp <- c("<is><t>V1</t></is>", "<is><t>V2</t></is>", "<is><t>N/A</t></is>")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$is)
  expect_equal(exp, got)


  set.seed(123)
  df <- data.frame(C = rnorm(10), D = rnorm(10))

  wb <- wb_workbook()$
    add_worksheet("df")$
    add_data(x = df, startCol = "C")
  # TODO add_formula()
  write_formula(wb, "df", startCol = "E", startRow = 2,
                x = "SUM(C2:C11*D2:D11)",
                array = TRUE)
  write_formula(wb, "df", x = "C3 + D3", startCol = "E", startRow = 3)
  x <- c(google = "https://www.google.com")
  class(x) <- "hyperlink"
  wb$add_data(sheet = "df", x = x, startCol = "E", startRow = 4)


  exp <- structure(
    list(c_t = c("", "str", "str"),
         f = c("SUM(C2:C11*D2:D11)", "C3 + D3", "=HYPERLINK(\"https://www.google.com\")"),
         f_t = c("array", "", "")),
    row.names = c("23", "110", "111"), class = "data.frame")
  got <- wb$worksheets[[1]]$sheet_data$cc[c(5, 8, 11), c("c_t", "f", "f_t")]
  expect_equal(exp, got)

  ### write logical
  xlsxFile <- testfile_path("readTest.xlsx")
  wb1 <- wb_load(xlsxFile)

  data <- head(wb_to_df(wb1, sheet = 3))
  wb <- wb_workbook()$add_worksheet()$add_data(x = data)$add_data(x = data)

  exp <- c("inlineStr", "", "b", "e")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$c_t)
  expect_equal(exp, got)

})

test_that("write dims", {

  # create a workbook
  wb <- wb_workbook()$
    add_worksheet()$add_data(dims = "B2:C3", x = matrix(1:4, 2, 2), colNames = FALSE)$
    add_worksheet()$add_data_table(dims = "B:C", x = as.data.frame(matrix(1:4, 2, 2)))$
    add_worksheet()$add_formula(dims = "B3", x = "42")

  s1 <- wb_to_df(wb, 1, colNames = FALSE)
  s2 <- wb_to_df(wb, 2, colNames = FALSE)
  s3 <- wb_to_df(wb, 3, colNames = FALSE)

  expect_equal(rownames(s1), c("2", "3"))
  expect_equal(rownames(s2), c("1", "2", "3"))
  expect_equal(rownames(s3), c("3"))

  expect_equal(colnames(s1), c("B", "C"))
  expect_equal(colnames(s2), c("B", "C"))
  expect_equal(colnames(s3), c("B"))

})

test_that("update cell(s)", {

  xlsxFile <- testfile_path("update_test.xlsx")
  wb <- wb_load(xlsxFile)

  # update Cells D4:D6 with 1:3
  wb <- wb_add_data(x = c(1:3),
                    wb = wb, sheet = "Sheet1", dims = "D4:D6")

  # update Cells B3:D3 (names())
  wb <- wb_add_data(x = c("x", "y", "z"),
                    wb = wb, sheet = "Sheet1", dims = "B3:D3")

  # update D4 again (single value this time)
  wb <- wb_add_data(x = 7,
                    wb = wb, sheet = "Sheet1", dims = "D4")

  # add new column on the left of the existing workbook
  wb <- wb_add_data(x = 7,
                    wb = wb, sheet = "Sheet1", dims = "A4")

  # add new row on the end of the existing workbook
  wb <- wb_add_data(x = 7,
                    wb = wb, sheet = "Sheet1", dims = "A9")

  exp <- structure(
    list(c(7, NA, NA, NA, NA, 7),
         c(NA, NA, TRUE, FALSE, TRUE, NA),
         c(2, NA, 2.5, NA, NA, NA),
         c(7, 2, 3, NA, 5, NA)),
    names = c(NA, "x", "Var2", "Var3"),
    row.names = 4:9,
    class = "data.frame")

  got <- wb_to_df(wb)
  expect_equal(exp, got, ignore_attr = TRUE)

  ####
  wb <- wb_workbook()$
    add_worksheet()$
    add_fill(dims = "B2:G8", color = wb_colour("yellow"))$
    add_data(dims = "C3", x = Sys.Date())$
    add_data(dims = "E4", x = Sys.Date(), removeCellStyle = TRUE)
  exp <- structure(list(r = c("B2", "C2", "D2", "E2", "F2", "G2"),
                        row_r = c("2", "2", "2", "2", "2", "2"),
                        c_r = c("B", "C", "D", "E", "F", "G"),
                        c_s = c("1", "1", "1", "1", "1", "1"),
                        c_t = c("", "", "", "", "", ""),
                        c_cm = c("", "", "", "", "", ""),
                        c_ph = c("", "", "", "", "", ""),
                        c_vm = c("", "", "", "", "", ""),
                        v = c("", "", "", "", "", ""),
                        f = c("", "", "", "", "", ""),
                        f_t = c("", "", "", "", "", ""),
                        f_ref = c("", "", "", "", "", ""),
                        f_ca = c("", "", "", "", "", ""),
                        f_si = c("", "", "", "", "", ""),
                        is = c("", "", "", "", "", ""),
                        typ = c("4", "4", "4", "4", "4", "4")),
                   row.names = 1:6,
                   class = "data.frame")
  got <- head(wb$worksheets[[1]]$sheet_data$cc)
  expect_equal(exp, got)

})

test_that("write_rownames", {
  wb <- wb_workbook()$
    add_worksheet()$add_data(x = mtcars, rowNames = TRUE)$
    add_worksheet()$add_data_table(x = mtcars, rowNames = TRUE)

  exp <- structure(
    list(A = c(NA, "Mazda RX4"), B = c("mpg", "21")),
    row.names = 1:2,
    class = "data.frame",
    tt = structure(
      list(A = c(NA, "s"), B = c("s", "n")),
      row.names = 1:2,
      class = "data.frame"),
    types = c(A = 0, B = 0)
  )
  got <- wb_to_df(wb, 1, dims = "A1:B2", colNames = FALSE, keep_attributes = TRUE)
  expect_equal(exp, got)

  exp <- structure(
    list(A = c("_rowNames_", "Mazda RX4"), B = c("mpg", "21")),
    row.names = 1:2,
    class = "data.frame",
    tt = structure(
      list(A = c("s", "s"), B = c("s", "n")),
      row.names = 1:2,
      class = "data.frame"),
    types = c(A = 0, B = 0)
  )
  got <- wb_to_df(wb, 2, dims = "A1:B2", colNames = FALSE, keep_attributes = TRUE)
  expect_equal(exp, got)

})

test_that("NA works as expected", {

  wb <- wb_workbook()$
    add_worksheet("Sheet1")$
    add_data(
      dims = "A1",
      x = NA,
      na.strings = NULL
    )$
    add_data(
      dims = "A2",
      x = NA_character_,
      na.strings = NULL
    )

  exp <- c(NA_real_, NA_real_)
  got <- wb_to_df(wb, colNames = FALSE)$A
  expect_equal(exp, got)

})

test_that("writeData() forces evaluation of x (#264)", {

  x <- format(123.4)
  df <- data.frame(d = format(123.4))
  df2 <- data.frame(e = x)

  wb <- wb_workbook()
  wb$add_worksheet("sheet")
  wb$add_data(startCol = 1, x = data.frame(a = format(123.4)))
  wb$add_data(startCol = 2, x = data.frame(b = as.character(123.4)))
  wb$add_data(startCol = 3, x = data.frame(c = "123.4"))
  wb$add_data(startCol = 4, x = df)
  wb$add_data(startCol = 5, x = df2)

  exp <- c(
    "<is><t>a</t></is>", "<is><t>b</t></is>", "<is><t>c</t></is>",
    "<is><t>d</t></is>", "<is><t>e</t></is>", "<is><t>123.4</t></is>"
  )
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$is)
  expect_equal(exp, got)

})

test_that("write character numerics with a correct cell style", {

  ## current default
  op <- options("openxlsx2.string_nums" = 0)
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_data(x = c("One", "2", "Three", "1.7976931348623157E+309", "2.5"))

  exp <- NA_character_
  got <- wb$styles_mgr$styles$cellXfs[2]
  expect_equal(exp, got)

  exp <- c("4", "4", "4", "4", "4")
  got <- wb$worksheets[[1]]$sheet_data$cc$typ
  expect_equal(exp, got)

  ## string numerics correctly flagged
  options("openxlsx2.string_nums" = 1)

  wb <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_data(x = c("One", "2", "Three", "1.7976931348623157E+309", "2.5")) %>%
    wb_add_worksheet() %>%
    wb_add_data(dims = "A1", x = "1992") %>%
    wb_add_data(dims = "A2", x = 1992) %>%
    wb_add_data(dims = "A3", x = "1992.a") %>%
    wb_add_worksheet() %>%
    wb_add_data(dims = "A1", x = 1e5) %>%
    wb_add_data(dims = "A2", x = "1e5") %>%
    wb_add_data(dims = "A3", x = 1e+05) %>%
    wb_add_data(dims = "A4", x = "1e+05")

  exp <- "<xf applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"49\" quotePrefix=\"1\" xfId=\"0\"/>"
  got <- wb$styles_mgr$styles$cellXfs[2]
  expect_equal(exp, got)

  exp <- c("4", "13", "4", "4", "13")
  got <- wb$worksheets[[1]]$sheet_data$cc$typ
  expect_equal(exp, got)

  exp <- c("13", "2", "4")
  got <- wb$worksheets[[2]]$sheet_data$cc$typ
  expect_equal(exp, got)

  exp <- c("2", "13", "2", "13")
  got <- wb$worksheets[[3]]$sheet_data$cc$typ
  expect_equal(exp, got)

  ## write string numerics as numerics (on the fly conversion)
  options("openxlsx2.string_nums" = 2)

  wb <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_data(x = c("One", "2", "Three", "1.7976931348623157E+309", "2.5"))

  exp <- NA_character_
  got <- wb$styles_mgr$styles$cellXfs[2]
  expect_equal(exp, got)

  exp <- c("4", "2", "4", "4", "2")
  got <- wb$worksheets[[1]]$sheet_data$cc$typ
  expect_equal(exp, got)
})

test_that("writing as shared string works", {

  df <- data.frame(
    x = letters,
    y = letters,
    stringsAsFactors = FALSE
  )

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = letters, dims = "A1", inline_strings = FALSE)$
    add_data(x = letters, dims = "B1", inline_strings = FALSE)$
    add_worksheet()$
    add_data(x = letters, dims = "A1", inline_strings = TRUE)$
    add_data(x = letters, dims = "B1", inline_strings = TRUE)$
    add_worksheet()$
    add_data_table(x = df, inline_strings = FALSE)$
    add_worksheet()$
    add_data_table(x = df, inline_strings = TRUE)

  expect_equal(letters, wb_to_df(wb, colNames = FALSE)$A)
  expect_equal(wb_to_df(wb, 1), wb_to_df(wb, 2))
  expect_equal(df, wb_to_df(wb, 3), ignore_attr = TRUE)
  expect_equal(wb_to_df(wb, 3), wb_to_df(wb, 4))

  expect_true(all(wb$worksheets[[1]]$sheet_data$cc$c_t == "s"))
  expect_true(all(wb$worksheets[[2]]$sheet_data$cc$c_t == "inlineStr"))

  # test missing cases in characters
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = c("a", NA, "b", "NA"), dims = "A1", inline_strings = FALSE)$
    add_worksheet()$
    add_data(x = c("a", NA, "b", "NA"), dims = "A1", inline_strings = FALSE, na.strings = "N/A")$
    add_worksheet()$
    add_data(x = c("a", NA, "b", "NA"), dims = "A1", inline_strings = FALSE, na.strings = NULL)

  exp <- structure(
    list(c_t = "e", v = "#N/A"),
    row.names = 2L,
    class = "data.frame"
  )
  got <- wb$worksheets[[1]]$sheet_data$cc[2, c("c_t", "v")]
  expect_equal(exp, got)

  exp <- structure(
    list(c_t = "s", v = "3"),
    row.names = 2L,
    class = "data.frame"
  )
  got <- wb$worksheets[[2]]$sheet_data$cc[2, c("c_t", "v")]
  expect_equal(exp, got)

  exp <- structure(
    list(c_t = "", v = ""),
    row.names = 2L,
    class = "data.frame"
  )
  got <- wb$worksheets[[3]]$sheet_data$cc[2, c("c_t", "v")]
  expect_equal(exp, got)

  # test missing cases in numerics
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = c(1L, NA, NaN, Inf), dims = "A1", inline_strings = FALSE)$
    add_worksheet()$
    add_data(x = c(1L, NA, NaN, Inf), dims = "A1", inline_strings = FALSE, na.strings = "N/A")$
    add_worksheet()$
    add_data(x = c(1L, NA, NaN, Inf), dims = "A1", inline_strings = FALSE, na.strings = NULL)

  expect_equal(wb_to_df(wb, 1), wb_to_df(wb, 3))
  expect_equal("N/A", wb_to_df(wb, 2)[1, 1])

})

test_that("writing pivot tables works", {

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)

  df <- wb_data(wb)

  wb$add_pivot_table(df, dims = "A3", filter = "am", rows = "cyl", cols = "gear", data = "hp")
  wb$add_pivot_table(df, dims = "A10", sheet = 2, rows = "cyl", cols = "gear", data = c("disp", "hp"), fun = "count")
  wb$add_pivot_table(df, dims = "A20", sheet = 2, rows = "cyl", cols = "gear", data = c("disp", "hp"), fun = "average")
  wb$add_pivot_table(df, dims = "A30", sheet = 2, rows = "cyl", cols = "gear", data = c("disp", "hp"), fun = c("sum", "average"))

  expect_equal(4L, length(wb$pivotTables))

})

test_that("writing pivot with escaped characters works", {

  example_df <- data.frame(
    location = c("London", "NYC", "NYC", "Berlin", "Madrid", "London", "Austin & Dallas"),
    amount = c(7, 5, 3, 2.5, 6, 1, 17)
  )

  wb <- wb_workbook() %>% wb_add_worksheet() %>% wb_add_data(x = example_df)
  df <- wb_data(wb)
  wb <- wb %>% wb_add_pivot_table(df, dims = "A3", rows = "location", data = "amount")

  cf <- xml_node(wb$pivotDefinitions, "pivotCacheDefinition", "cacheFields", "cacheField")[1]

  exp <- "<s v=\"Austin &amp; Dallas\"/>"
  got <- xml_node(cf, "cacheField", "sharedItems", "s")[5]
  expect_equal(exp, got)

})

test_that("writing slicers works", {

  wb <- wb_workbook() %>%
    ### Sheet 1
    wb_add_worksheet() %>%
    wb_add_data(x = mtcars)

  df <- wb_data(wb, sheet = 1)

  varname <- c("vs", "drat")

  ### Sheet 2
  wb$
    # first pivot
    add_pivot_table(
      df, dims = "A3", slicer = varname, rows = "cyl", cols = "gear", data = "disp",
      pivot_table = "mtcars"
    )$
    add_slicer(x = df, sheet = current_sheet(), slicer = "vs", pivot_table = "mtcars")$
    add_slicer(x = df, dims = "B18:D24", sheet = current_sheet(), slicer = "drat", pivot_table = "mtcars",
               params = list(columnCount = 5))$
    # second pivot
    add_pivot_table(
      df, dims = "G3", sheet = current_sheet(), slicer = varname, rows = "gear", cols = "carb", data = "mpg",
      pivot_table = "mtcars2"
    )$
    add_slicer(x = df, dims = "G12:I16", slicer = "vs", pivot_table = "mtcars2",
               params = list(sortOrder = "descending", caption = "Wow!"))

  ### Sheet 3
  wb$
    add_pivot_table(
      df, dims = "A3", slicer = varname, rows = "gear", cols = "carb", data = "mpg",
      pivot_table = "mtcars3"
    )$
    add_slicer(x = df, dims = "A12:D16", slicer = "vs", pivot_table = "mtcars3")

  # test a few conditions
  expect_equal(2L, length(wb$slicers))
  expect_equal(4L, length(wb$slicerCaches))
  expect_equal(xml_node_name(wb$workbook$extLst, "extLst", "ext"), "x14:slicerCaches")
  expect_equal(1L, wb$worksheets[[2]]$relships$slicer)
  expect_equal(2L, wb$worksheets[[3]]$relships$slicer)
  expect_equal(25L, grep("slicer2.xml", wb$Content_Types))

  ## test error
  wb <- wb_workbook() %>%
  wb_add_worksheet() %>%
    wb_add_data(x = mtcars)

  df <- wb_data(wb, sheet = 1)

  varname <- c("vs", "drat")

  wb$
    # first pivot
    add_pivot_table(
      df, dims = "A3", rows = "cyl", cols = "gear", data = "disp",
      params = list(name = "mtcars")
    )

  expect_error(
    wb$add_slicer(x = df, sheet = current_sheet(), slicer = "vs", pivot_table = "mtcars"),
    "slicer was not initialized in pivot table!"
  )

})

test_that("writing na.strings = NULL works", {

  # write na.strings = na_strings()
  tmp <- temp_xlsx()
  write_xlsx(matrix(NA, 2, 2), tmp)
  wb <- wb_load(tmp)

  exp <- "#N/A"
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$v[3:6])
  expect_equal(exp, got)

  # write na.strings = ""
  tmp <- temp_xlsx()
  write_xlsx(matrix(NA, 2, 2), tmp, na.strings = "")
  wb <- wb_load(tmp)

  exp <- "<is><t/></is>"
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$is[3:6])
  expect_equal(exp, got)

  # write na.strings = NULL
  tmp <- temp_xlsx()
  write_xlsx(matrix(NA, 2, 2), tmp, na.strings = NULL)
  wb <- wb_load(tmp)

  exp <- ""
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$v[3:6])
  expect_equal(exp, got)

  got <- unique(wb$worksheets[[1]]$sheet_data$cc$is[3:6])
  expect_equal(exp, got)

})

# write third party data.frame classes
test_that("write data.table class", {

  dt <- structure(
    list(
      a_number = c(1234, 4321),
      a_string = c("hello", "world"),
      a_date = structure(c(19358, 19448), class = "Date"),
      a_boolean = c(FALSE, TRUE)
    ),
    class = c("data.table", "data.frame"),
    row.names = c(NA, -2L)
  )

  tmp <- temp_xlsx()
  expect_silent(write_xlsx(dt, tmp))
  expect_equal(dt, read_xlsx(tmp), ignore_attr = TRUE)

})

test_that("write tibble class", {

  tbl <- structure(
    list(
      a_number = c(1234, 4321),
      a_string = c("hello", "world"),
      a_date = structure(c(19358, 19448), class = "Date"),
      a_boolean = c(FALSE, TRUE)
    ),
    class = c("tbl_df", "tbl", "data.frame"),
    row.names = c(NA, -2L)
  )

  tmp <- temp_xlsx()
  expect_silent(write_xlsx(tbl, tmp))
  expect_equal(tbl, read_xlsx(tmp), ignore_attr = TRUE)

})


test_that("writing labeled variables works", {

  x <- c(1, 2, 1, -99, -97)
  attr(x, "labels") <- c("N/A" = -97, "NaN" = -98, "NA" = -99)

  exp <- c("1", "2", "1", "NA", "N/A")
  got <- to_string(x)

  wb <- wb_workbook()$add_worksheet()$add_data(x = x)

  exp <- c("1", "2", "1",
           "<is><t>x</t></is>", "<is><t>NA</t></is>", "<is><t>N/A</t></is>")
  cc <- wb$worksheets[[1]]$sheet_data$cc[c("v", "is")]
  cc[cc$v == "", "v"] <- NA
  cc[cc$is == "", "is"] <- NA
  got <- unlist(cc[!is.na(cc)])
  expect_equal(exp, got)

  x <- factor(x = c("M", "F"), levels = c("M", "F"), labels = c(1L, 2L))
  exp <- c("1", "2")
  got <- to_string(x)
  expect_equal(exp, got)

  wb <- wb_workbook()$add_worksheet()$add_data(x = x)
  exp <- c(1, 2)
  got <- wb_to_df(wb, colNames = FALSE)$A
  expect_equal(exp, got)

})

test_that("writing in specific encoding works", {

  skip_on_cran()

  loc <- Sys.getlocale("LC_CTYPE")
  Sys.setlocale("LC_CTYPE", "")

  op <- options(
    "openxlsx2.force_utf8_encoding" = TRUE,
    "openxlsx2.native_encoding" = "CP1251"
  )
  on.exit(options(op), add = TRUE)

  # a cyrillic string: https://github.com/JanMarvin/openxlsx2/issues/640
  enc_str <- as.raw(c(0xd0, 0xb0, 0xd0, 0xb1, 0xd0, 0xb2, 0xd0, 0xb3, 0xd0,
                      0xb4))
  enc_str <- rawToChar(enc_str)
  Encoding(enc_str) <- "UTF-8"

  loc_str <- stringi::stri_encode(enc_str, from = "UTF-8", to = "CP1251")

  tmp <- temp_xlsx()
  wb <- wb_workbook()$add_worksheet("sheet")$add_data("sheet", x = loc_str)
  wb$save(tmp)
  expect_silent(wb2 <- wb_load(tmp))

  # exp <- wb$worksheets[[1]]$sheet_data$cc$is[1]
  # got <- wb2$worksheets[[1]]$sheet_data$cc$is[1]
  # expect_equal(exp, got)

  # got <- stringi::stri_encode(wb_to_df(wb, colNames = FALSE)$A, from = "UTF-8", to = "CP1251")
  # expect_equal(enc_str, got)

  tmp <- tempfile()
  write_file(head = "<a>", body = loc_str, tail = "</a>", fl = tmp)
  # got <- xml_value(tmp, "a")
  exp <- loc_str
  got <- stringi::stri_encode(xml_value(tmp, "a"), from = "UTF-8", to = "CP1251")
  expect_equal(exp, got)

  Sys.setlocale("LC_CTYPE", loc)

})

test_that("writing NULL works silently", {

  tmp <- temp_xlsx()
  x <- NULL

  expect_silent(write_xlsx(x, tmp))

  expect_silent(wb_workbook()$add_worksheet()$add_data(x = x))

  wb <- wb_workbook()$add_worksheet()
  wb2 <- wb_add_data(wb, x = x)
  expect_equal(wb, wb2)

})
