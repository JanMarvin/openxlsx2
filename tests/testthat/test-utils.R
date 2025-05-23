
test_that("paste_c() works", {
  res <- paste_c("a", "", "b", NULL, "c", " ")
  exp <- "a b c  "
  expect_identical(res, exp)

  expect_identical(paste_c(character(), "", NULL), "")
})

test_that("rbindlist", {

  expect_equal(data.frame(), rbindlist(character()))

})

test_that("dims to col & row and back", {

  exp <- list(col = c("A", "B"), row = c("1", "2"))
  got <- dims_to_rowcol("A1:B2")
  expect_equal(exp, got)

  exp <- list(col = 1:2, row = 1:2)
  got <- dims_to_rowcol("A1:B2", as_integer = TRUE)
  expect_equal(exp, got)

  exp <- list(col = 1:2, row = c(1L))
  got <- dims_to_rowcol("A:B", as_integer = TRUE)
  expect_equal(exp, got)

  exp <- list(col = "A", row = c("1"))
  got <- dims_to_rowcol("A:A", as_integer = FALSE)
  expect_equal(exp, got)

  exp <- "A1:A1"
  got <- rowcol_to_dims(1, "A")
  expect_equal(exp, got)

  exp <- "A1:A10"
  got <- rowcol_to_dims(1:10, 1)
  expect_equal(exp, got)

  exp <- "E2:J8"
  got <- rowcol_to_dims(2:8, 5:10)
  expect_equal(exp, got)

  # We might have to relax this in the future, if we are ever going to allow
  # something like "$1:$2" to select every column for a range of cells.
  expect_error(dims_to_rowcol("A1;B64:65"), "A dims string passed to ")

  exp <- "A1:A3,B1:B3"
  got <- rowcol_to_dims(row = c(1, 3), col = c(1, 2), single = FALSE)
  expect_equal(exp, got)

})

test_that("`wb_dims()` works/errors as expected with unnamed arguments", {

  # Acceptable inputs
  expect_error(wb_dims())
  expect_error(wb_dims(NULL))
  expect_equal(wb_dims(1L, 1L), "A1")
  expect_equal(wb_dims(1:10, 1:26), "A1:Z10")
  expect_equal(wb_dims(1:10, LETTERS), "A1:Z10")
  expect_equal(
    wb_dims(1:10, 1:12, from_row = 2),
    wb_dims(rows = 1:10, cols = 1:12, from_row = 2)
  )
  expect_warning(wb_dims("1", 2), "should not be a character")
  expect_error(wb_dims(from_row = "C"), "from_col/from_row must be positive integers if supplied.")

  # Ambiguous / input not accepted.
  # This now fails, as it used not to work. (Use `wb_dims()`, `NULL`, or )
  expect_error(wb_dims(1), "Supplying a single unnamed argument.")

  # This used to return A1 as well.
  expect_error(wb_dims(2), "Supplying a single unnamed argument is not handled")
  expect_error(wb_dims(mtcars), "Supplying a single unnamed argument")


  expect_error(wb_dims(rows = c(1), cols = c(-1)), "You must supply positive values to `cols`")
  expect_error(wb_dims(rows = c(-1), cols = c(1)), "You must supply positive values to `rows`")

  expect_error(wb_dims(1, 2, 3))
})

test_that("`wb_dims()` errors when providing unsupported arguments", {
  expect_error(
    wb_dims(cols = 1:10, col = 5:7),
    "invalid argument"
  )
  expect_error(
    wb_dims(rows = 1:10, ffrom_col = 5:7),
    "invalid argument"
  )
  expect_error(wb_dims(rows = 1:10, start_row = 5), "`from_row`")
  expect_error(wb_dims(start_col = 2), "`from_col`")
  # providing a vector to `from_row` or `from_col`
  expect_error(wb_dims(from_row = 5:7))
  expect_error(wb_dims(fom_col = 5:7))
})

test_that("wb_dims() works when not supplying `x`.", {
  expect_equal(wb_dims(rows = 1:10, cols = 5:7), "E1:G10")
  expect_equal(wb_dims(rows = 5:7, cols = 1:10), "A5:J7")
  expect_equal(wb_dims(rows = 5, cols = 7), "G5")

  expect_equal(wb_dims(1:2, 1:4, from_row = 2, from_col = "B"), "B2:E3")
  # Assuming other unnamed argument is rows or cols.
  expect_message(out <- wb_dims(1, rows = 2), "Assuming the .+ `cols`")
  expect_equal(out, "A2")
  expect_message(out <- wb_dims(cols = 1, 2), "Assuming the .+ `rows`")
  expect_equal(out, "A2")
  # warns when trying to pass weird things
  expect_error(wb_dims(rows = "BC", cols = 1), "`rows` is character and contains nothing that can be interpreted as number.")
  # "`wb_dims()` works
  expect_equal(wb_dims(from_col = 4), "D1")
  expect_equal(wb_dims(from_row = 4), "A4")
  expect_equal(wb_dims(from_row = 4, from_col = 3), "C4")
  expect_equal(wb_dims(from_row = 4, from_col = "C"), "C4")

  expect_equal(wb_dims(4, 3), wb_dims(from_row = 4, from_col = "C"))
  expect_error(wb_dims(0, 3))
  expect_error(wb_dims(3, 0))
  expect_error(wb_dims(1, 1, col_names = TRUE))
  expect_error(wb_dims(1, 1, row_names = FALSE), "`row_names`")
})

test_that("wb_dims() works with above and one line", {
  expect_equal(wb_dims(x = names(mtcars), from_row = 2), "A2:A12")

  # doesn't work
  # expect_equal(wb_dims(x = names(mtcars), from_row = 3, above = 1), "A2:A12")

  expect_warning(
    expect_equal(
      wb_dims(x = names(mtcars), from_row = 2, above = 1),
      "A1:A11"
    ),
    "rows cannot be above of row 1 "
  )

  expect_warning(
    expect_equal(
      wb_dims(x = names(mtcars), from_dims = "A2", above = 1),
      "A1:A11"),
    "rows cannot be above of row 1 "
  )

  # Doesn't work returns A1:A11
  # expect_equal(wb_dims(x = names(mtcars), from_dims = "A3", above = 1), "A2:A12")
  # Doesn't work returns A1:A11
  # expect_equal(wb_dims(x = names(mtcars), from_dims = "A4", above = 1), "A3:A13")

  # currently the only ways to create horizontal single row dims with wb_dims()

  nms <- mtcars[NULL, ]

  expect_equal(wb_dims(x = nms, from_row = 3), "A3:K3")
  expect_equal(wb_dims(x = nms, from_row = 3, above = 1), "A2:K2")
  expect_equal(wb_dims(x = nms, from_dims = "A3"), "A3:K3")
  expect_equal(wb_dims(x = nms, from_dims = "A3", above = 1), "A2:K2")
  expect_equal(wb_dims(x = nms, from_dims = "A4", above = 1), "A3:K3")

  expect_equal(wb_dims(x = t(names(mtcars)), col_names = FALSE, from_row = 3), "A3:K3")
  expect_equal(wb_dims(x = t(names(mtcars)), col_names = FALSE, from_row = 3, above = 1), "A2:K2")
  expect_equal(wb_dims(x = t(names(mtcars)), col_names = FALSE, from_dims = "A3"), "A3:K3")
  expect_equal(wb_dims(x = t(names(mtcars)), col_names = FALSE, from_dims = "A3", above = 1), "A2:K2")
  expect_equal(wb_dims(x = t(names(mtcars)), col_names = FALSE, from_dims = "A4", above = 1), "A3:K3")

})

test_that("`wb_dims()` can select content in a nice fashion with `x`", {
  # Selecting content
  # Assuming that the data was written to a workbook with:
  # col_names = TRUE, start_col = "B", start_row = 2, row_names = FALSE
  wb_dims_cars <- function(...) {
    wb_dims(x = mtcars, from_row = 2, from_col = "B", ...)
  }
  full_data_dims <- wb_dims_cars()
  expect_equal(full_data_dims, "B2:L34")
  # selecting only content (data)
  data_content_dims <- "B3:L34"
  expect_equal(wb_dims_cars(select = "data"), data_content_dims)
  # Selecting column names
  col_names_dims <- "B2:L2"
  expect_equal(wb_dims_cars(select = "col_names"), col_names_dims)
  expect_equal(
    wb_dims_cars(select = "col_names"),
    wb_dims(rows = 1, cols = seq_len(ncol(mtcars)), from_row = 2, from_col = "B")
  )

  # Selecting a column "cyl"
  dims_cyl <- "C3:C34"
  expect_equal(wb_dims_cars(cols = "cyl"), dims_cyl)
  expect_equal(wb_dims_cars(cols = 2), dims_cyl)


  # Supplying a column range
  dims_col1_2 <- "B3:C34"
  expect_equal(wb_dims_cars(cols = 1:2), dims_col1_2)

  # Supplying a column range, but select column names too
  dims_col1_2_with_name <- "B2:C34"
  expect_equal(wb_dims_cars(cols = 1:2, select = "x"), dims_col1_2_with_name)


  # Selecting a row range
  dims_row1_to_5 <- "B3:L7"
  expect_equal(wb_dims(x = mtcars, from_row = 2, from_col = "B", rows = 1:5), dims_row1_to_5)

  # Select a row range with the names of `x`
  dims_row1_to_5_and_names <- "B2:L7"
  expect_equal(wb_dims(x = mtcars, from_row = 2, from_col = "B", rows = 1:5, select = "x"), dims_row1_to_5_and_names)
})

test_that("wb_dims can select multiple columns and rows", {

  exp <- "B2:B33"
  got <- wb_dims(x = mtcars, cols = c("cyl"))
  expect_equal(exp, got)

  exp <- "B2:B33,J2:J33"
  got <- wb_dims(x = mtcars, cols = c("cyl", "gear"))
  expect_equal(exp, got)

  exp <- "B1,J1"
  got <- wb_dims(x = mtcars, cols = c("cyl", "gear"), select = "col_names")
  expect_equal(exp, got)

  exp <- "A1:K33"
  got <- wb_dims(x = mtcars)
  expect_equal(exp, got)

  exp <- "B2:B33"
  got <- wb_dims(x = mtcars, cols = c(2))
  expect_equal(exp, got)

  exp <- "B2:B33,D2:D33"
  got <- wb_dims(x = mtcars, cols = c(2, 4))
  expect_equal(exp, got)

  exp <- "A5:K5"
  got <- wb_dims(x = mtcars, rows = c(4))
  expect_equal(exp, got)

  exp <- "A3:K3,A5:K5,A6:K6"
  got <- wb_dims(x = mtcars, rows = c(2, 4:5))
  expect_equal(exp, got)

  exp <- "B3,D3,B5,D5,B6,D6"
  got <- wb_dims(x = mtcars, cols = c(2, 4), rows = c(2, 4:5))
  expect_equal(exp, got)

})

test_that("`wb_dims()` works for a matrix without column names", {
  mt <- matrix(c(1, 2))
  expect_equal(wb_dims(x = mt), "A1:A3")
  expect_equal(wb_dims(x = mt, select = "data"), "A2:A3")
  expect_equal(wb_dims(x = mt, col_names = FALSE), "A1:A2")
  expect_equal(wb_dims(x = mt, row_names = TRUE, col_names = TRUE), "A1:B3")
  expect_equal(wb_dims(x = mt, select = "col_names"), "A1")
})

test_that("`wb_dims()` works when Supplying an object `x`.", {
  expect_equal(wb_dims(x = mtcars, col_names = TRUE), "A1:K33")
  expect_equal(wb_dims(x = mtcars), "A1:K33")
  expect_equal(wb_dims(x = mtcars, select = "data"), "A2:K33")
  out <- wb_dims(x = mtcars, col_names = FALSE)
  expect_equal(out, "A1:K32")
  # doesn't work
  expect_equal(wb_dims(x = letters), "A1:A26")
  expect_equal(wb_dims(x = t(letters), col_names = TRUE), "A1:Z2")
  expect_error(wb_dims(x = letters, col_names = TRUE), "Can't supply `col_names` when `x` is a vector") # don't want this error anymore.

  expect_equal(wb_dims(x = mtcars, rows = 5, from_col = "C"), "C6:M6")

  expect_equal(wb_dims(x = mtcars, from_row = 2, from_col = "B"), "B2:L34")
  expect_equal(wb_dims(x = mtcars, from_row = 2, from_col = "B", col_names = FALSE), "B2:L33")

  expect_equal(wb_dims(x = mtcars, rows = 5:10, from_col = "C"), "C6:M11")
  # Write without column names on top

  expect_error(wb_dims(x = mtcars, cols = 0, from_col = "C"), "You must supply positive values to `cols`")

  # select rows and columns work
  expect_equal(wb_dims(x = mtcars, rows = 2:10, cols = "cyl"), "B3:B11")

  expect_equal(wb_dims(rows = 1 + seq_len(nrow(mtcars)), cols = 4), "D2:D33")
  out_hp <- wb_dims(x = mtcars, cols = "hp") # , "col name = 'hp' to `cols = 4`")
  expect_equal(out_hp, "D2:D33")
  expect_equal(out_hp, wb_dims(rows = 1 + seq_len(nrow(mtcars)), cols = 4))
  # select column name also

  out_hp_with_cnam <- wb_dims(x = mtcars, cols = "hp", select = "x") # , message =  "col name = 'hp' to `cols = 4`")
  expect_equal(out_hp_with_cnam, "D1:D33")
  expect_equal(out_hp_with_cnam, wb_dims(rows = 1:(nrow(mtcars) + 1), cols = 4))

  expect_equal(wb_dims(x = mtcars, cols = 4, select = "x"), "D1:D33")
  expect_equal(wb_dims(x = mtcars, from_col = 2, select = "data"), "B2:L33")

  # use 1 column name works

  expect_error(wb_dims(cols = "hp"), "Supplying a single argument")
  # using non-existing character column doesn't work
  expect_error(wb_dims(x = mtcars, cols = "A"), "`cols` must be an integer or an existing column name of `x`.")
  expect_equal(wb_dims(x = mtcars, cols = c("hp", "vs")), "D2:D33,H2:H33")
  expect_error(expect_warning(wb_dims(x = mtcars, rows = "hp")), "`rows` is character and contains nothing that can be interpreted as number.")
  # Access only row / col name
  expect_no_message(wb_dims(x = mtcars, select = "col_names"))
  # to write without column names, specify `from_row = 0` (or -1 of what you wanted)
})

test_that("`wb_dims()` handles row_names = TRUE consistenly.", {
  # Select the data grid when row names are present
  expect_equal(wb_dims(x = mtcars, row_names = TRUE), "A1:L33")

  # select row names (with the top left corner cell)
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, col_names = TRUE, select = "row_names"), "A2:A33")
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, col_names = FALSE, select = "row_names"), "A1:A32")

  # select x + column names (without rows)
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, col_names = TRUE, select = "data"), "B2:L33")

  # an object without column names and row names works.
  expect_equal(wb_dims(x = unname(mtcars), row_names = TRUE, col_names = FALSE), "B1:L32")

  # Selecting rows is also correct

  # column positions are still respected with row names
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, cols = "cyl"), "C2:C33")
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, cols = 2:4), "C2:E33")
  # select row names only
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, select = "row_names", from_col = "B"), "B2:B33")
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, select = "row_names", from_row = 2), "A3:A34")
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, from_col = "B", from_row = 2, select = "data"), "C3:M34")

  # selecting both rows and columns doesn't work
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, rows = 2:10, cols = "cyl"), "C3:C11")
  # Select the data + row names
  expect_warning(out <- wb_dims(x = mtcars, row_names = TRUE, select = "x", from_row = "2"), "supply an integer")
  expect_equal(out, "A2:L34")
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, col_names = FALSE, from_row = 2, select = "x"), "A2:L33") # col_span would need to be col_span+1 in this case.
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, col_names = FALSE, from_row = 2, select = "data"), "B2:L33") # col_span would need to be col_span+1 in this case.

  # Selecting the full grid with row names + col names is a bit more complex
  expect_equal(wb_dims(x = mtcars, row_names = TRUE, col_names = TRUE), "A1:L33")
  expect_equal(wb_dims(x = mtcars, rows = 2:10, cols = "cyl", row_names = TRUE), "C3:C11")
  # Style row names of an object
})

test_that("wb_dims() errors clearly with bad `select`.", {
  expect_error(wb_dims(x = mtcars, select = c("bad", "col_names")), "accepts a single")
  expect_error(wb_dims(2, 10, select = "col_names"), "Can't supply `select` when `x` is absent")

  expect_error(wb_dims(x = mtcars, col_names = FALSE, select = "col_names"), "col_names = TRUE")
  expect_error(wb_dims(x = mtcars, col_names = FALSE, select = "row_names"), "row_names = TRUE")
  expect_error(wb_dims(x = mtcars, row_names = FALSE, select = "row_names"), "row_names = TRUE")
  # row_names = FALSE is the default
  expect_error(wb_dims(x = mtcars, select = "row_names"), "row_names = TRUE")
})

test_that("wb_dims() handles empty data frames", {
  expect_equal(wb_dims(x = data.frame(), from_col = "B"), "B1")
})

test_that("wb_dims() handles `from_dims`", {
  expect_equal(
    wb_dims(from_dims = "A3"),
    "A3"
  )
  expect_error(
    wb_dims(from_dims = "A1", from_col = 2)
  )
  expect_error(
    wb_dims(from_dims = "A1", from_row = 2)
  )
  expect_error(
    wb_dims(from_dims = "A1", from_row = 2, from_col = 2)
  )
  expect_equal(
    wb_dims(x = mtcars, from_dims = "B7"),
    wb_dims(x = mtcars, from_col = "B", from_row = 7)
  )
  expect_error(wb_dims(from_dims = "65"))
})

test_that("wb_dims() corner cases work", {

  exp <- "B2,B30"
  got <- wb_dims(rows = c(2, 30), cols = 2)       # expect B2,B30
  expect_equal(exp, got)

  exp <- "B1,B30"
  got <- wb_dims(rows = c(1, 30), cols = 2)       # expect B1,B30
  expect_equal(exp, got)

  exp <- "B11,B40"
  got <- wb_dims(rows = c(1, 30), cols = 2, from_row = 11)
  expect_equal(exp, got)

  exp <- "B2,D2,B30,D30"
  got <- wb_dims(rows = c(2, 30), cols = c(2, 4)) # expect B2,D2,B30,D30
  expect_equal(exp, got)

  exp <- "B1,D1,B30,D30"
  got <- wb_dims(rows = c(1, 30), cols = c(2, 4)) # expect B1,D1,B30,D30
  expect_equal(exp, got)

  exp <- "B2:B1,D2:D1"
  got <- wb_dims(rows = c(2, 1), cols = c(2, 4))  # expect B2,D2,B1,D1
  expect_equal(exp, got)

  exp <- "B5,D5,B1,D1"
  got <- wb_dims(rows = c(5, 1), cols = c(2, 4))  # expect B5,D5,B1,D1
  expect_equal(exp, got)

})

test_that("row and col selection works", {

  # equal with sample cols
  expect_equal(
    wb_dims(x = mtcars, cols = c(1, 4:7, 9)),
    wb_dims(rows = seq_len(nrow(mtcars)) + 1L, cols = c(1, 4:7, 9))
  )

  # unequeal with sample rows
  expect_equal(
    wb_dims(x = mtcars, rows = which(mtcars$cyl == 6)),
    wb_dims(rows = which(mtcars$cyl == 6) + 1L, cols = seq_along(mtcars))
  )

  expect_equal(
    wb_dims(x = mtcars, rows = which(mtcars$cyl == 6), from_row = 4),
    wb_dims(rows = which(mtcars$cyl == 6) + 1L, cols = seq_along(mtcars), from_row = 4)
  )

  expect_equal(
    wb_dims(x = mtcars, rows = which(mtcars$cyl == 6), from_col = 6),
    wb_dims(rows = which(mtcars$cyl == 6) + 1L, cols = seq_along(mtcars), from_col = 6)
  )

  expect_equal(
    wb_dims(x = mtcars, cols = c(1, 4:7, 9), row_names = TRUE),
    wb_dims(rows = seq_len(nrow(mtcars)) + 1L, cols = c(1, 4:7, 9), from_col = 2)
  )

})

test_that("create_char_dataframe", {

  exp <- data.frame(x1 = rep("", 5), z1 = rep("", 5), stringsAsFactors = FALSE)

  got <- create_char_dataframe(colnames = c("x1", "z1"), n = 5)

  expect_equal(exp, got)

})

test_that("test random_string", {
  set.seed(123)
  options("openxlsx2_seed" = NULL)

  x <- .Random.seed
  tmp <- random_string()
  y <- .Random.seed
  expect_identical(x, y)
  expect_equal(tmp, "HmPsw2WtYSxSgZ6t")

  x <- .Random.seed
  tmp <- random_string(length = 6)
  y <- .Random.seed
  expect_identical(x, y)
  expect_equal(tmp, "GNZuCt")

  x <- .Random.seed
  tmp <- random_string(length = 6, keep_seed = FALSE)
  y <- .Random.seed
  expect_false(identical(x, y))
})

test_that("temp_dir works", {

  out <- temp_dir("temp_dir_works")
  expect_true(dir.exists(out))

})

test_that("relship", {

  wb <- wb_load(testfile_path("loadExample.xlsx"))

  exp <- "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing3.xml\"/>"
  got <- relship_no(obj = wb$worksheets_rels[[3]], x = "table")
  expect_equal(got, exp)

  got <- get_relship_id(obj = wb$worksheets_rels[[1]], x = "drawing")
  expect_equal(got, "rId2")

})

test_that("as_xml_attr works", {

  mm <- matrix("", 2, 2)
  # as_xml_attr is internal
  mm[1, 1] <- as_xml_attr(TRUE)
  mm[2, 1] <- as_xml_attr(FALSE)

  mm[1, 2] <- as_xml_attr(1)
  mm[2, 2] <- as_xml_attr(NULL)

  exp <- matrix(c("1", "0", "1", ""), 2, 2)
  expect_equal(exp, mm)

})

test_that("string formating", {

  foo <- fmt_txt("foo: ", bold = TRUE, size = 16, color = wb_color("green"))
  bar <- fmt_txt("bar", underline = TRUE)
  txt <- paste0(foo, bar)

  exp <- "<is><r><rPr><b/><sz val=\"16\"/><color rgb=\"FF00FF00\"/></rPr><t xml:space=\"preserve\">foo: </t></r></is>"
  got <- txt_to_is(foo)
  expect_equal(exp, got)

  exp <- "<si><r><rPr><b/><sz val=\"16\"/><color rgb=\"FF00FF00\"/></rPr><t xml:space=\"preserve\">foo: </t></r></si>"
  got <- txt_to_si(foo)
  expect_equal(exp, got)

  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(txt))

  exp <- "foo: bar"
  got <- wb_to_df(wb)$txt
  expect_equal(exp, got)

  exp <- structure(
    "<r><rPr><i/><strike/><rFont val=\"Arial\"/><charset/><outline/><vertAlign val=\"baseline\"/></rPr><t>test</t></r>",
    class = c("character", "fmt_txt")
  )
  got <- fmt_txt(
    "test",
    italic = TRUE,
    strike = TRUE,
    font = "Arial",
    charset = "",
    outline = TRUE,
    vert_align = "baseline"
  )
  expect_equal(exp, got)

})

test_that("outdec = \",\" works", {

  op <- options(OutDec = ",")
  on.exit(options(op), add = TRUE)

  exp <- "[1] 1,1"
  got <- capture.output(print(1.1))
  expect_equal(exp, got)

  wb <- wb_workbook()$add_worksheet()$set_col_widths(
    cols = c(1, 4, 6, 7, 9),
    widths = c(16, 15, 12, 18, 33)
  )

  exp <- c("16.711", "8.43", "15.711", "8.43", "12.711", "18.711", "8.43", "33.711")
  got <- rbindlist(xml_attr(wb$worksheets[[1]]$cols_attr, "col"))$width
  expect_equal(exp, got)

  exp <- "[1] 1,1"
  got <- capture.output(print(1.1))
  expect_equal(exp, got)

})

test_that("fmt_txt works", {

  exp <- structure(
    "<r><rPr><b/></rPr><t xml:space=\"preserve\">Foo </t></r>",
    class = c("character", "fmt_txt")
  )
  got <- fmt_txt("Foo ", bold = TRUE)
  expect_equal(exp, got)

  exp <- structure(
    "<r><rPr><b/></rPr><t xml:space=\"preserve\">Hello </t></r><r><rPr/><t xml:space=\"preserve\">World </t></r>",
    class = c("character", "fmt_txt")
  )
  got <- fmt_txt("Hello ", bold = TRUE) + fmt_txt("World ")
  expect_equal(exp, got)

  expect_message(got <- capture.output(print(got)), "fmt_txt string:")
  expect_equal(got, "[1] \"Hello World \"")

  ## watch out!
  txt <- fmt_txt("Sum ", bold = TRUE) + 2 + 2
  exp <- "Sum 22"
  got <- as.character(txt)
  expect_equal(exp, got)

  txt <- fmt_txt("Sum ", bold = TRUE) + (2 + 2)
  exp <- "Sum 4"
  got <- as.character(txt)
  expect_equal(exp, got)

  txt <- fmt_txt("0 < 1", bold = TRUE)
  exp <- "0 < 1"
  got <- as.character(txt)
  expect_equal(exp, got)

  exp <- "<r><rPr><b/></rPr><t>0 &lt; 1</t></r>"
  got <- unclass(txt)
  expect_equal(exp, got)

  expect_equal(
    fmt_txt("foo", color = wb_color("green")),
    fmt_txt("foo", colour = wb_colour("green"))
  )

})

test_that("protect against blank string in fmt_txt()", {
  exp <- "<r><rPr/><t/></r><r><rPr/><t>a+b</t></r>"
  got <- unclass(fmt_txt(x = "", size = 12) + fmt_txt(x = "a+b"))
  expect_equal(exp, got)
})

test_that("fmt_txt() input validation works", {

  exp <- "<r><rPr><b/><i/><u val=\"double\"/><strike/><sz val=\"12\"/><color rgb=\"FF0000FF\"/><rFont val=\"Times New Roman\"/><charset val=\"1\"/><outline/><vertAlign val=\"baseline\"/><family val=\"2\"/><shadow/><condense/><extend/></rPr><t>foo bar</t></r>"
  fmt_str <- fmt_txt(
    x = "foo bar",
    bold = TRUE,
    italic = TRUE,
    strike = TRUE,
    size = 12,
    color = wb_color("blue"),
    font = "Times New Roman",
    charset = 1,
    outline = TRUE,
    family = 2,
    shadow = TRUE,
    condense = TRUE,
    vert_align = "baseline",
    underline = "double",
    extend = TRUE
  )
  got <- unclass(fmt_str)
  expect_equal(exp, got)

  expect_error(
    fmt_txt(
      x = "foo bar",
      color = "blue",
      font = "Times New Roman",
      condense = 0,
      vert_align = "top"
    ),
    "color must be class wbColour"
  )

  expect_error(
    fmt_txt(
      x = "foo bar",
      font = "Times New Roman",
      condense = 0,
      vert_align = "top"
    ),
    "Use any of 'baseline', 'superscript', 'subscript'"
  )

})

test_that("wb_dims(from_dims) positioning works", {

  mm <- matrix(1:4, 2, 2)

  # positioning
  dims1 <- wb_dims(from_dims = "B2")
  dims2 <- wb_dims(from_dims = dims1, below = 2)
  dims3 <- wb_dims(from_dims = dims2, right = 2)
  dims4 <- wb_dims(from_dims = dims3, above = 2)
  dims5 <- wb_dims(from_dims = dims4, left = 2)

  exp <- c("B2", "B4", "D4", "D2", "B2")
  got <- c(dims1, dims2, dims3, dims4, dims5)
  expect_equal(exp, got)

  # positioning with x
  dims1 <- wb_dims(from_dims = "B2", x = mm)
  dims2 <- wb_dims(from_dims = dims1, x = mm, below = 2)
  dims3 <- wb_dims(from_dims = dims2, x = mm, right = 2)
  dims4 <- wb_dims(from_dims = dims3, x = mm, above = 2)
  dims5 <- wb_dims(from_dims = dims4, x = mm, left = 2)

  exp <- c("B2:C4", "B6:C8", "E6:F8", "E2:F4", "B2:C4")
  got <- c(dims1, dims2, dims3, dims4, dims5)
  expect_equal(exp, got)

  # col_names = FALSE
  dims1 <- wb_dims(from_dims = "B2", x = mm, col_names = FALSE)
  dims2 <- wb_dims(from_dims = dims1, x = mm, below = 2, col_names = FALSE)
  dims3 <- wb_dims(from_dims = dims2, x = mm, right = 2, col_names = FALSE)
  dims4 <- wb_dims(from_dims = dims3, x = mm, above = 2, col_names = FALSE)
  dims5 <- wb_dims(from_dims = dims4, x = mm, left = 2, col_names = FALSE)

  exp <- c("B2:C3", "B5:C6", "E5:F6", "E2:F3", "B2:C3")
  got <- c(dims1, dims2, dims3, dims4, dims5)
  expect_equal(exp, got)

  # row_names = TRUE
  dims1 <- wb_dims(from_dims = "B2", x = mm, row_names = TRUE)
  dims2 <- wb_dims(from_dims = dims1, x = mm, below = 2, row_names = TRUE)
  dims3 <- wb_dims(from_dims = dims2, x = mm, right = 2, row_names = TRUE)
  dims4 <- wb_dims(from_dims = dims3, x = mm, above = 2, row_names = TRUE)
  dims5 <- wb_dims(from_dims = dims4, x = mm, left = 2, row_names = TRUE)

  exp <- c("B2:D4", "B6:D8", "F6:H8", "F2:H4", "B2:D4")
  got <- c(dims1, dims2, dims3, dims4, dims5)
  expect_equal(exp, got)

})

test_that("dims are quick", {

  # full range
  ddims <- dims_to_dataframe(
    "A1:D5000",
    fill = TRUE
  )
  edims <- dims_to_dataframe(
    "A1:A5000,B1:B5000,C1:C5000,D1:D5000",
    fill = TRUE
  )
  expect_equal(ddims, edims)

  # with blank in between
  ddims <- dims_to_dataframe(
    "A1:B5000,D1:E5000",
    fill = TRUE
  )
  edims <- dims_to_dataframe(
    "A1:A5000,B1:B5000,D1:D5000,E1:E5000",
    fill = TRUE
  )
  expect_equal(ddims, edims)

  # with blank in between and cleaned up
  ddims <- dims_to_dataframe(
    "A1:B5000,D1:E5000",
    fill = TRUE, empty_rm = TRUE
  )
  edims <- dims_to_dataframe(
    "A1:A5000,B1:B5000,D1:D5000,E1:E5000",
    fill = TRUE, empty_rm = TRUE
  )
  expect_equal(ddims, edims)

  expect_silent(dims_to_dataframe("A:C"))
  expect_silent(dims_to_dataframe("3:3"))
  expect_error(dims_to_dataframe("A:3"))

})

test_that("wb_dims does not try to validate column names", {

  df <- data.frame(
    a1 = 1,
    aa2 = 2,
    aaa3 = 3,
    aaaa4 = 4,
    aaaaa5 = 5,
    aaaaaa6 = 6,
    aaaaaaa7 = 7,
    aaaaaaaa8 = 8,
    aaaaaaaaa9 = 9
  )

  exp <- paste0(LETTERS[1:9], "2")
  got <- vector("character", length = ncol(df))
  for (i in seq_along(df)) {
    got[i] <- wb_dims(x = df, cols = colnames(df)[i], rows = 1)
  }
  expect_equal(exp, got)

})

test_that("wb_dims handles decreasing consecutive columns and rows", {

  exp <- "A3:A1,B3:B1,C3:C1,H3:H1,I3:I1,J3:J1"
  got <- wb_dims(rows = 3:1, cols = c(1:3, 8:10))
  expect_equal(exp, got)

  exp <- "A3:A1,B3:B1,C3:C1,J3:J1,I3:I1,H3:H1"
  got <- wb_dims(rows = 3:1, cols = c(1:3, 10:8))
  expect_equal(exp, got)

  exp <- "C3:C1,B3:B1,A3:A1,J3:J1,I3:I1,H3:H1"
  got <- wb_dims(rows = 3:1, cols = c(3:1, 10:8))
  expect_equal(exp, got)

})

test_that("dims are the same", {
  blank  <- dims_to_dataframe("C3:A1", fill = FALSE)
  filled <- dims_to_dataframe("C3:A1", fill = TRUE)
  expect_equal(dimnames(blank), dimnames(filled))

  one <- dims_to_dataframe("A1:C3", fill = TRUE)
  two <- dims_to_dataframe("C3:A1", fill = TRUE)

  expect_equal(
    unname(unlist(one)),
    rev(unname(unlist(two)))
  )

  # this is a probably exclusively internal handler that can be used to create
  # non consecutive dims for wb_add_fill(). At the moment it does not appear to
  # be required to keep the column and row ordering of each of the pieces.
  # This might also be non trivial for something like "B3:A1,A5:C6"
  # (is it up/down and left/right?)
  exp <- dims_to_dataframe("B1:C2,B5:C8", fill = FALSE)
  got <- dims_to_dataframe("C1:B2,C5:B8", fill = FALSE)
  expect_equal(exp, got)

})

test_that("reverse rows and cols work", {
  expect_equal("A1:C3", wb_dims(rows = 1:3, cols = 1:3))
  expect_equal("C3:A1", wb_dims(rows = 3:1, cols = 3:1))
  expect_equal("C1:A3", wb_dims(rows = 1:3, cols = 3:1))
  expect_equal("A3:C1", wb_dims(rows = 3:1, cols = 1:3))
})
