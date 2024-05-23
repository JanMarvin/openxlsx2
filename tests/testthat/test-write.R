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
  write_formula(wb, "df", start_col = "E", start_row = 2,
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)

  cc <- wb$worksheets[[1]]$sheet_data$cc
  got <- cc[cc$row_r == "2" & cc$c_r == "E", ]
  expect_equal(exp[1:16], got[1:16])


  rownames(exp) <- 1L
  # write formula first add data later
  wb <- wb_workbook()
  wb <- wb_add_worksheet(wb, "df")
  write_formula(wb, "df", start_col = "E", start_row = 2,
               x = "SUM(C2:C11*D2:D11)",
               array = TRUE)
  wb$add_data("df", df, start_col = "C")

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
                    x = mtcars, dims = "B3", row_names = TRUE,
                    table_style = "TableStyleLight9")

  # [1:4] to ignore factor
  expect_equal(iris[1:4], wb_to_df(wb, "S1")[1:4], ignore_attr = TRUE)
  expect_equal(iris[1:4], wb_to_df(wb, "S1")[1:4], ignore_attr = TRUE)

  # handle rownames
  got <- wb_to_df(wb, "S2", row_names = TRUE)
  attr(got, "tt") <- NULL
  attr(got, "types") <- NULL
  expect_equal(got, mtcars)
  expect_equal(rownames(got), rownames(mtcars))

})

test_that("test options", {

  ops <- options()
  tmp <- temp_xlsx()
  wb_workbook()$add_worksheet("Sheet 1")$add_data("Sheet 1", cars)
  ops2 <- options()

  # adding data to the worksheet should not alter the global options
  expect_equal(ops, ops2)

})

test_that("missing x is caught early [#246]", {
  expect_error(
    wb_workbook()$add_data(mtcars),
    "`x` is missing"
  )
  expect_error(
    wb_workbook()$add_data_table(mtcars),
    "`x` is missing"
  )
})

test_that("missing sheet is caught early (#942)", {
  expect_error(
    wb_workbook()$add_data(x = mtcars),
    "no worksheet"
  )
  expect_error(
    wb_workbook()$add_data_table(x = mtcars),
    "no worksheet"
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
    add_data(x = df, start_col = "C")
  wb$add_formula("df", start_col = "E", start_row = 2,
                x = "SUM(C2:C11*D2:D11)",
                array = TRUE)
  wb$add_formula("df", x = "C3 + D3", start_col = "E", start_row = 3)
  x <- c(google = "https://www.google.com")
  class(x) <- "hyperlink"
  wb$add_data(sheet = "df", x = x, start_col = "E", start_row = 4)


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
    add_worksheet()$add_data(dims = "B2:C3", x = matrix(1:4, 2, 2), col_names = FALSE)$
    add_worksheet()$add_data_table(dims = "B:C", x = as.data.frame(matrix(1:4, 2, 2)))$
    add_worksheet()$add_formula(dims = "B3", x = "42")

  s1 <- wb_to_df(wb, 1, col_names = FALSE)
  s2 <- wb_to_df(wb, 2, col_names = FALSE)
  s3 <- wb_to_df(wb, 3, col_names = FALSE)

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
    names = c(NA, "x", "y", "z"),
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
    add_worksheet()$add_data(x = mtcars, row_names = TRUE)$
    add_worksheet()$add_data_table(x = mtcars, row_names = TRUE)

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
  got <- wb_to_df(wb, 1, dims = "A1:B2", col_names = FALSE, keep_attributes = TRUE)
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
  got <- wb_to_df(wb, 2, dims = "A1:B2", col_names = FALSE, keep_attributes = TRUE)
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
  got <- wb_to_df(wb, col_names = FALSE)$A
  expect_equal(exp, got)

})

test_that("writeData() forces evaluation of x (#264)", {

  x <- format(123.4)
  df <- data.frame(d = format(123.4))
  df2 <- data.frame(e = x)

  wb <- wb_workbook()
  wb$add_worksheet("sheet")
  wb$add_data(start_col = 1, x = data.frame(a = format(123.4)))
  wb$add_data(start_col = 2, x = data.frame(b = as.character(123.4)))
  wb$add_data(start_col = 3, x = data.frame(c = "123.4"))
  wb$add_data(start_col = 4, x = df)
  wb$add_data(start_col = 5, x = df2)

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

  got <- wb$styles_mgr$styles$cellXfs[2]
  expect_equal(got, NA_character_)

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

  got <- wb$styles_mgr$styles$cellXfs[2]
  expect_equal(got, NA_character_)

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

  expect_equal(letters, wb_to_df(wb, col_names = FALSE)$A)
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
  expect_equal(wb_to_df(wb, 2)[1, 1], "N/A")

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

  expect_equal(length(wb$pivotTables), 4L)

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
  expect_length(wb$slicers, 2L)
  expect_length(wb$slicerCaches, 4L)
  expect_equal(xml_node_name(wb$workbook$extLst, "extLst", "ext"), "x14:slicerCaches")
  expect_equal(wb$worksheets[[2]]$relships$slicer, 1L)
  expect_equal(wb$worksheets[[3]]$relships$slicer, 2L)
  expect_equal(grep("slicer2.xml", wb$Content_Types), 25L)

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

test_that("writing slicers works", {

  dat <- data.frame(
    date = seq(from = as.Date("2024-01-01"), length.out = 26, by = "month"),
    amnt = sample(seq(100:150), 26, replace = TRUE),
    lttr = letters[1:2]
  )

  wb <- wb_workbook()$add_worksheet()$add_data(x = dat)
  df <- wb_data(wb)
  wb$add_pivot_table(x = df, cols = "lttr", data = "amnt", timeline = "date", pivot_table = "pivot1")
  wb$add_timeline(x = df, timeline = "date", pivot_table = "pivot1")

  expect_equal("x15:timelineCacheRefs", xml_node_name(wb$workbook$extLst, "extLst", "ext"))
  expect_equal("timelines", xml_node_name(wb$timelines))
  expect_equal("timelineCacheDefinition", xml_node_name(wb$timelineCaches))
  expect_true(grepl("x15:timelineRefs", wb$worksheets[[2]]$extLst))

  wb$add_slicer(df, slicer = "lttr", pivot_table = "pivot1")

  expect_equal(c("x15:timelineCacheRefs", "x14:slicerCaches"), xml_node_name(wb$workbook$extLst, "extLst", "ext"))
  expect_equal(c("x15:timelineRefs", "x14:slicerList"), xml_node_name(wb$worksheets[[2]]$extLst, "ext"))

  # and the other way around it works too
  wb <- wb_workbook()$add_worksheet()$add_data(x = dat)
  df <- wb_data(wb)
  wb$add_pivot_table(x = df, cols = "lttr", data = "amnt", timeline = "date", pivot_table = "pivot1")
  wb$add_slicer(df, slicer = "lttr", pivot_table = "pivot1")
  wb$add_timeline(x = df, timeline = "date", pivot_table = "pivot1")

  expect_equal(c("x14:slicerCaches", "x15:timelineCacheRefs"), xml_node_name(wb$workbook$extLst, "extLst", "ext"))
  expect_equal(c("x14:slicerList", "x15:timelineRefs"), xml_node_name(wb$worksheets[[2]]$extLst, "ext"))

})


test_that("removing slicers works", {

  ### prepare data
  df <- data.frame(
    AirPassengers = c(AirPassengers),
    time = seq(from = as.Date("1949-01-01"), to = as.Date("1960-12-01"), by = "month"),
    letters = letters[1:4]
  )

  ### create workbook
  wb <- wb_workbook()$
    add_worksheet("pivot")$
    add_worksheet("pivot2")$
    add_worksheet("data")$
    add_data(x = df)

  ### get pivot table data source
  df <- wb_data(wb, sheet = "data")

  ### first sheet
  # create pivot table
  wb$add_pivot_table(
    df,
    sheet = "pivot",
    rows = "time",
    cols = "letters",
    data = "AirPassengers",
    pivot_table = "airpassengers",
    params = list(
      compact = FALSE, outline = FALSE, compact_data = FALSE,
      row_grand_totals = FALSE, col_grand_totals = FALSE)
  )

  # add slicer
  wb$add_slicer(
    df,
    dims = "E1:I7",
    sheet = "pivot",
    slicer = "letters",
    pivot_table = "airpassengers",
    params = list(choose = c(letters = 'x %in% c("a", "b")'))
  )

  wb$add_slicer(
    df,
    dims = "E8:I15",
    sheet = "pivot",
    slicer = "time",
    pivot_table = "airpassengers"
  )

  ### second sheet
  # create pivot table
  wb$add_pivot_table(
    df,
    sheet = "pivot2",
    rows = "time",
    cols = "letters",
    data = "AirPassengers",
    pivot_table = "airpassengers2",
    params = list(
      compact = FALSE, outline = FALSE, compact_data = FALSE,
      row_grand_totals = FALSE, col_grand_totals = FALSE)
  )

  # add slicer
  wb$add_slicer(
    df,
    dims = "E1:I7",
    sheet = "pivot2",
    slicer = "letters",
    pivot_table = "airpassengers2",
    params = list(choose = c(letters = 'x %in% c("a", "b")'))
  )

  wb$add_slicer(
    df,
    dims = "E8:I15",
    sheet = "pivot2",
    slicer = "time",
    pivot_table = "airpassengers2"
  )

  ### remove slicer
  wb$remove_slicer(sheet = "pivot")

  temp <- temp_xlsx()
  expect_silent(wb$save(temp)) # no warning, all files written as expected

})

test_that("removing timelines works", {

  ### prepare data
  df <- data.frame(
    AirPassengers = c(AirPassengers),
    time = seq(from = as.Date("1949-01-01"), to = as.Date("1960-12-01"), by = "month"),
    letters = letters[1:4]
  )

  ### create workbook
  wb <- wb_workbook()$
    add_worksheet("pivot")$
    add_worksheet("pivot2")$
    add_worksheet("data")$
    add_data(x = df)

  ### get pivot table data source
  df <- wb_data(wb, sheet = "data")

  ### first sheet
  # create pivot table
  wb$add_pivot_table(
    df,
    sheet = "pivot",
    rows = "time",
    cols = "letters",
    data = "AirPassengers",
    pivot_table = "airpassengers",
    params = list(
      compact = FALSE, outline = FALSE, compact_data = FALSE,
      row_grand_totals = FALSE, col_grand_totals = FALSE)
  )

  # add slicer
  wb$add_slicer(
    df,
    dims = "E1:I7",
    sheet = "pivot",
    slicer = "letters",
    pivot_table = "airpassengers"
  )

  # add timeline
  wb$add_timeline(
    df,
    dims = "E9:I14",
    sheet = "pivot",
    timeline = "time",
    pivot_table = "airpassengers"
  )

  ### second sheet
  # create pivot table
  wb$add_pivot_table(
    df,
    sheet = "pivot2",
    rows = "time",
    cols = "letters",
    data = "AirPassengers",
    pivot_table = "airpassengers2",
    params = list(
      compact = FALSE, outline = FALSE, compact_data = FALSE,
      row_grand_totals = FALSE, col_grand_totals = FALSE)
  )

  # add slicer
  wb$add_slicer(
    df,
    dims = "E1:I7",
    sheet = "pivot2",
    slicer = "letters",
    pivot_table = "airpassengers2",
    params = list(choose = c(letters = 'x %in% c("a", "b")'))
  )

  # add timeline
  wb$add_timeline(
    df,
    dims = "E9:I14",
    sheet = "pivot2",
    timeline = "time",
    pivot_table = "airpassengers2"
  )

  ### remove slicer
  wb$remove_timeline(sheet = "pivot")

  temp <- temp_xlsx()
  expect_silent(wb$save(temp)) # no warning, all files written as expected

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
  got <- wb_to_df(wb, col_names = FALSE)$A
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

  # got <- stringi::stri_encode(wb_to_df(wb, col_names = FALSE)$A, from = "UTF-8", to = "CP1251")
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

test_that("dimension limits work", {

  max_c <- 16384
  max_r <- 1048576

  dims <- paste0(int2col(max_c), max_r)
  expect_silent(
    wb <- wb_workbook()$add_worksheet()$add_data(x = 1, dims = dims)
  )

  dims <- paste0(int2col(max_c), max_r + 1L)
  expect_error(
    wb_workbook()$add_worksheet()$add_data(x = 1, dims = dims),
    "Dimensions exceed worksheet"
  )

  dims <- paste0(int2col(max_c + 1L), max_r)
  expect_error(
    wb_workbook()$add_worksheet()$add_data(x = 1, dims = dims),
    "Dimensions exceed worksheet"
  )

  dims <- paste0(int2col(max_c + 1L), max_r + 1L)
  expect_error(
    wb_workbook()$add_worksheet()$add_data(x = 1, dims = dims),
    "Dimensions exceed worksheet"
  )

})

test_that("numfmt option works", {

  op <- options("openxlsx2.numFmt" = "$ #.0")
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()$add_worksheet()$add_data(x = 1:10)

  exp <- "<numFmt numFmtId=\"165\" formatCode=\"$ #.0\"/>"
  got <- wb$styles_mgr$styles$numFmts
  expect_equal(exp, got)

})

test_that("comma option works", {

  op <- options("openxlsx2.commaFormat" = "#.0")
  on.exit(options(op), add = TRUE)

  dat <- data.frame(x = 1:10 + rnorm(1:10))
  class(dat$x) <- c("comma", class(dat$x))

  wb <- wb_workbook()$add_worksheet()$add_data(x = dat)

  exp <- "<numFmt numFmtId=\"165\" formatCode=\"#.0\"/>"
  got <- wb$styles_mgr$styles$numFmts
  expect_equal(exp, got)

})

test_that("filter works with wb_add_data()", {

  wb <- wb_workbook()$
    add_worksheet()$add_data(x = mtcars, with_filter = TRUE)$
    add_worksheet()$add_data(x = mtcars, with_filter = TRUE)$
    add_data(x = cars, with_filter = TRUE)

  exp <- "<autoFilter ref=\"A1:K33\"/>"
  got <- wb$worksheets[[1]]$autoFilter
  expect_equal(exp, got)

  exp <- c(
    "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">'Sheet 1'!$A$1:$K$33</definedName>",
    "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"1\" hidden=\"1\">'Sheet 2'!$A$1:$B$51</definedName>"
  )
  got <- wb$workbook$definedNames
  expect_equal(exp, got)

})

test_that("writing total row works", {

  # default row sums
  wb <- wb_workbook()$add_worksheet()$add_data_table(x = mtcars, total_row = TRUE)

  exp <- data.frame(
    A = "SUBTOTAL(109,Table1[mpg])", B = "SUBTOTAL(109,Table1[cyl])",
    C = "SUBTOTAL(109,Table1[disp])", D = "SUBTOTAL(109,Table1[hp])",
    E = "SUBTOTAL(109,Table1[drat])", F = "SUBTOTAL(109,Table1[wt])",
    G = "SUBTOTAL(109,Table1[qsec])", H = "SUBTOTAL(109,Table1[vs])",
    I = "SUBTOTAL(109,Table1[am])", J = "SUBTOTAL(109,Table1[gear])",
    K = "SUBTOTAL(109,Table1[carb])"
  )
  got <- wb_to_df(wb, dims = wb_dims(rows = 33, cols = "A:K"),
                  show_formula = TRUE, col_names = FALSE)
  expect_equal(exp, got, ignore_attr = TRUE)

  # empty total row
  wb <- wb_workbook()$add_worksheet()$add_data_table(x = mtcars, total_row = c("none"))

  exp <- data.frame(
    A = NA_real_, B = NA_real_, C = NA_real_, D = NA_real_,
    E = NA_real_, F = NA_real_, G = NA_real_, H = NA_real_, I = NA_real_,
    J = NA_real_, K = NA_real_
  )
  got <- wb_to_df(wb, dims = wb_dims(rows = 33, cols = "A:K"),
                  show_formula = TRUE, col_names = FALSE)
  expect_equal(exp, got, ignore_attr = TRUE)

  # total row with text only
  wb <- wb_workbook()$add_worksheet()$add_data_table(x = cars, total_row = c(text = "Result", text = "sum"))

  exp <- data.frame(A = "Result", B = "sum")
  got <- wb_to_df(wb, dims = wb_dims(rows = 51, cols = "A:B"),
                  show_formula = TRUE, col_names = FALSE)
  expect_equal(exp, got, ignore_attr = TRUE)

  # total row with text and formula
  wb <- wb_workbook()$add_worksheet()$add_data_table(x = cars, total_row = c(text = "Result", "sum"))

  exp <- data.frame(A = "Result", B = "SUBTOTAL(109,Table1[dist])")
  got <- wb_to_df(wb, dims = wb_dims(rows = 51, cols = "A:B"),
                  show_formula = TRUE, col_names = FALSE)
  expect_equal(exp, got, ignore_attr = TRUE)

  # total row with none and custom formula
  wb <- wb_workbook()$add_worksheet()$add_data_table(x = cars, total_row = c("none", "COUNTA"))

  exp <- data.frame(A = NA_real_, B = "COUNTA(Table1[dist])")
  got <- wb_to_df(wb, dims = wb_dims(rows = 51, cols = "A:B"),
                  show_formula = TRUE, col_names = FALSE)
  expect_equal(exp, got, ignore_attr = TRUE)

  # with rownames
  wb <- wb_workbook()$add_worksheet()$
    add_data_table(
      x = as.data.frame(USPersonalExpenditure),
      row_names = TRUE,
      total_row = c(text = "Total", "none", "sum", "sum", "sum", "SUM")
    )

  exp <- data.frame(
    A = "Total", B = NA_real_, C = "SUBTOTAL(109,Table1[1945])",
    D = "SUBTOTAL(109,Table1[1950])", E = "SUBTOTAL(109,Table1[1955])",
    F = "SUM(Table1[1960])"
  )
  got <- wb_to_df(wb, dims = wb_dims(rows = 6, cols = "A:F"), col_names = FALSE, show_formula = TRUE)
  expect_equal(exp, got, ignore_attr = TRUE)

})

test_that("writing vectors direction with dims works", {

  # write vectors column or rowwise
  wb <- wb_workbook()$add_worksheet()$
    add_data(x = 1:2, dims = "A1:B1")$
    add_data(x = t(1:2), dims = "D1:D2", col_names = FALSE)$ # ignores dims
    add_data(x = 1:2, dims = "A3:A4")$
    add_data(x = t(1:2), dims = "D3:E3", col_names = FALSE)  # ignores dims

  exp <- c("A1", "B1", "D1", "E1", "A3", "D3", "E3", "A4")
  got <- wb$worksheets[[1]]$sheet_data$cc$r
  expect_equal(exp, got)

  ## sum, sum as array and sum as cm
  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = head(cars))$
    add_formula(x = c("SUM(A2:A7)", "SUM(B2:B7)"), dims = "A9:B9")$
    add_formula(x = c("{SUM(A2:A7)}", "{SUM(B2:B7)}"), dims = "A10:B10")

  expect_warning(
    wb$add_formula(x = c("{SUM(A2:A7)}", "{SUM(B2:B7)}"), dims = "A11:B11", cm = TRUE),
    "modifications with cm formulas are experimental. use at own risk"
  )

  exp <- c("A1", "B1", "A2", "B2", "A3", "B3", "A4", "B4", "A5", "B5",
           "A6", "B6", "A7", "B7", "A9", "B9", "A10", "B10", "A11", "B11")
  got <- wb$worksheets[[1]]$sheet_data$cc$r
  expect_equal(exp, got)

})

test_that("dims size warnings work", {

  op <- options("openxlsx2.warn_if_dims_dont_fit" = TRUE)
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()$add_worksheet()

  # default no dims
  expect_warning(
    wb$add_data(x = head(mtcars)),
    "dimension of `x` exceeds all `dims`"
  )

  # with explicit default dims
  expect_warning(
    wb$add_data(dims = "A1", x = head(mtcars)),
    "dimension of `x` exceeds all `dims`"
  )

  # wb_add_data(dims = wb_dims(x = obj), x = obj) should always be silent
  expect_silent(wb$add_data(dims = wb_dims(x = head(mtcars)), x = head(mtcars)))

  # correct size should always be silent
  expect_silent(wb$add_data(dims = "A1:K7", x = head(mtcars)))

  # To wide
  expect_warning(
    wb$add_data(dims = "A1:K1", x = head(mtcars)),
    "dimension of `x` exceeds rows of `dims`"
  )

  # To short
  expect_warning(
    wb$add_data(dims = "A1:J7", x = head(mtcars)),
    "dimension of `x` exceeds cols of `dims`"
  )

  # ending in the correct cell isn't enough
  expect_warning(
    wb$add_data(dims = "B2:K7", x = head(mtcars)),
    "dimension of `x` exceeds all `dims`"
  )

  # currently write_xlsx() uses the default dims
  expect_warning(
    wb <- write_xlsx(x = head(mtcars))
  )

})

test_that("writing zero row data frames works", {

  # write an empty data frame
  dat <- data.frame()
  expect_silent(wb <- wb_workbook()$add_worksheet()$add_data(x = dat))

  exp <- NULL
  expect_message(got <- wb_to_df(wb), "sheet found, but contains no data")
  expect_equal(exp, got)

  # write a data frame containing an empty date vector
  dat <- data.frame(date = base::as.Date(NULL))
  expect_silent(wb <- wb_workbook()$add_worksheet()$add_data(x = dat))

  exp <- "date"
  got <- names(wb_to_df(wb))
  expect_equal(exp, got)

  # try the same with write_xlsx()
  expect_silent(wb <- write_xlsx(x = dat))

  exp <- "date"
  got <- names(wb_to_df(wb))
  expect_equal(exp, got)

})

test_that("non consecutive columns do not overlap", {
  test_dt <- data.frame(
    V1 = seq(as.Date("2024-01-01"), as.Date("2024-01-05"), 1),
    V2 = letters[1:5],
    V3 = seq(as.Date("2024-01-01"), as.Date("2024-01-05"), 1),
    V4 = c(letters[3:5], NA, NA),
    V5 = 1:5,
    V6 = c(NA, NA, 3, 4, 5),
    V7 = letters[1:5],
    V8 = c(letters[3:5], NA, NA),
    V9 = 1:5,
    V0 = c(NA, NA, 3, 4, 5)
  )

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = test_dt)

  # df <- wb_to_df(wb, col_names = F)

  cc <- wb$worksheets[[1]]$sheet_data$cc

  exp <- ""
  got <- cc[cc$r == "B2", "c_s"]
  expect_equal(exp, got)

})
