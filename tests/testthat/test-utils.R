
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

  exp <- list(c("A", "B"), c("1", "2"))
  got <- dims_to_rowcol("A1:B2")
  expect_equal(exp, got)

  exp <- list(1:2, 1:2)
  got <- dims_to_rowcol("A1:B2", as_integer = TRUE)
  expect_equal(exp, got)

  exp <- list(1:2, c(1L))
  got <- dims_to_rowcol("A:B", as_integer = TRUE)
  expect_equal(exp, got)

  exp <- list("A", c("1"))
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

})

test_that("`wb_dims()` works/errors as expected with unnamed arguments", {
  # Acceptable inputs
  expect_equal(wb_dims(), "A1")
  expect_equal(wb_dims(NULL), "A1") # to help programming with `wb_dims()` maybe?
  expect_equal(wb_dims(1L, 1L), "A1")
  expect_equal(wb_dims(1:10, 1:26), "A1:Z10")
  expect_equal(wb_dims(1:10, LETTERS), "A1:Z10")
  expect_equal(
    wb_dims(1:10, 1:12, from_row = 2),
    wb_dims(rows = 1:10, cols = 1:12, from_row = 2)
  )

  # Ambiguous / input not accepted.
  # This now fails, as it used not to work. (Use `wb_dims()`, `NULL`, or )
  expect_error(wb_dims(1), "Supplying a single unnamed argument.")
  # This used to return A1 as well.
  expect_error(wb_dims(2), "Supplying a single unnamed argument is not handled")
  expect_error(wb_dims(mtcars), "Supplying a single unnamed argument")
  # "`wb_dims()` WIP"
  skip("lower priority, but giving non-consecutive rows, or cols should error.")
  expect_error(wb_dims(rows = c(1, 3, 4), cols = c(1, 4)), "wb_dims() should only be used for Supplying a single continuous range.")
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
  # This used to error, but now passes with a message.
  expect_message(out <- wb_dims(1, rows = 2), "Assuming the .+ `cols`")
  expect_equal(out, "A2")
  # warns when trying to pass weird things
  expect_warning(wb_dims(rows = "BC", cols = 1), regexp = "integer.+`rows`")
  # "`wb_dims()` newe
  expect_equal(wb_dims(from_col = 4), "D1")
  expect_equal(wb_dims(from_row = 4), "A4")
  expect_equal(wb_dims(from_row = 4, from_col = 3), "C4")
  expect_equal(wb_dims(from_row = 4, from_col = "C"), "C4")

  expect_equal(wb_dims(4, 3), wb_dims(from_row = 4, from_col = "C"))
  expect_error(wb_dims(0, 3))
  expect_error(wb_dims(3, 0))
  expect_error(wb_dims(1, 1, col_names = TRUE))
  expect_error(wb_dims(1, 1, row_names = FALSE))

})
test_that("`wb_dims()` can select content in a nice fashion with `x`", {
  # Selecting content
  # Assuming that the data was written to a workbook with:
  # col_names = TRUE, start_col = "B", start_row = 2, row_names = FALSE
  wb_dims_cars <- function(...) {
    wb_dims(x = mtcars, from_row = 2, from_col = "B", ...)
  }
  full_data_dims <- wb_dims_cars(col_names = TRUE)
  expect_equal(full_data_dims, "B2:L34")

  # Selecting column names
  col_names_dims <- "B2:L2"
  expect_equal(wb_dims_cars(rows = 0), col_names_dims)
  expect_equal(
    wb_dims_cars(rows = 0),
    wb_dims(rows = 1, cols = seq_len(ncol(mtcars)), from_row = 2, from_col = "B")
  )
  # selecting only content (data)
  data_content_dims <- "B3:L34"
  expect_equal(wb_dims_cars(), data_content_dims)

  # Selecting a column "cyl"
  dims_cyl <- "C3:C34"
  expect_equal(suppressMessages(wb_dims_cars(cols = "cyl")), dims_cyl)
  expect_equal(suppressMessages(wb_dims_cars(cols = 2)), dims_cyl)


  # Supplying a column range
  dims_col1_2 <- "B3:C34"
  expect_equal(suppressMessages(wb_dims_cars(cols = 1:2)), dims_col1_2)

  # Supplying a column range, but select column names too
  dims_col1_2_with_name <- "B2:C34"
  expect_equal(wb_dims_cars(cols = 1:2, col_names = TRUE), dims_col1_2_with_name)


  # Selecting a row range
  dims_row1_to_5 <- "B3:L7"
  expect_equal(wb_dims_cars(rows = 1:5), dims_row1_to_5)

  # Select a row range with the names of `x`
  dims_row1_to_5_and_names <- "B2:L7"
  expect_equal(wb_dims_cars(rows = 0:5), dims_row1_to_5_and_names)
})

test_that("`wb_dims()` works for a matrix without column names", {
  mt <- matrix(c(1, 2))
  wb_dims(x = mt)
  wb_dims(x = mt, col_names = TRUE)
  wb_dims(x = mt, col_names = FALSE)
  expect_warning(dims_with_warning <- wb_dims(x = mtcars, col_names = FALSE), "`x` has column nam")
  expect_no_warning(dims_with_no_warning <- wb_dims(x = unname(mtcars), from_row = 1))
  expect_equal(dims_with_warning, dims_with_no_warning)
})

test_that("`wb_dims()` works when Supplying an object `x`.", {
  expect_equal(wb_dims(x = mtcars, col_names = TRUE), "A1:K33")
  expect_equal(wb_dims(x = mtcars), "A2:K33")
  expect_warning(out <- wb_dims(x = mtcars, col_names = FALSE), "`x` has column names")
  expect_equal(out, "A2:K33")


  expect_equal(wb_dims(x = letters), "A1:A26")

  expect_equal(wb_dims(x = t(letters), col_names = TRUE), "A1:Z2")
  expect_equal(wb_dims(x = mtcars, rows = 5, from_col = "C"), "C6:M6")

  expect_equal(wb_dims(x = mtcars, from_row = 2, from_col = "B", col_names = TRUE), "B2:L34")
  # previously
  expect_equal(wb_dims(x = mtcars), "A2:K33")
  expect_error(wb_dims(x = letters, col_names = TRUE), "Supplying `col_names` when `x` is a vector is not supported.")
  expect_equal(wb_dims(x = mtcars, rows = 5:10, from_col = "C"), "C6:M11")
  # Write without column names on top
  # select the full data `use if previously, you didn't write column name.
  expect_equal(wb_dims(x = mtcars), "A2:K33")
  # select the full data of an object without colnames work
  expect_equal(wb_dims(x = unname(mtcars), col_names = FALSE), "A1:K32")

  expect_error(wb_dims(x = mtcars, from_row = 0), "Use `rows = 0` to select column names")
  expect_error(wb_dims(x = mtcars, cols = 0, from_col = "C"), "`rows = 0`")
  expect_equal(wb_dims(x = mtcars, rows = 0), "A1:K1")
  expect_equal(wb_dims(x = mtcars, cols = 0, row_names = TRUE), "A2:A33")
  # If you want to include the first row as well.
  expect_equal(wb_dims(x = mtcars, cols = 0, row_names = TRUE, col_names = TRUE), "A1:A33")
  expect_equal(wb_dims(x = mtcars, cols = 0, row_names = TRUE), "A2:A33")
  expect_equal(wb_dims(x = mtcars, rows = 0, row_names = TRUE), "B1:L1")

  # expect_r

  expect_equal(wb_dims(rows = 1 + 1:nrow(mtcars), cols = 4), "D2:D33")
  expect_message(out_hp <- wb_dims(x = mtcars, cols = "hp"), "col name = 'hp' to `cols = 4`")
  expect_equal(out_hp, "D2:D33")
  expect_equal(out_hp, wb_dims(rows = 1 + 1:nrow(mtcars), cols = 4))
  # select column name also

  expect_message(out_hp_with_cnam <- wb_dims(x = mtcars, cols = "hp", col_names = TRUE), "col name = 'hp' to `cols = 4`")
  expect_equal(out_hp_with_cnam, "D1:D33")
  expect_equal(out_hp_with_cnam, wb_dims(rows = 1:(nrow(mtcars) + 1), cols = 4))

  expect_equal(wb_dims(x = mtcars, cols = 4, col_names = TRUE), "D1:D33")
  expect_error(wb_dims(x = mtcars, col_names = TRUE, from_row = 0), "Use `rows = 0`")
  expect_error(wb_dims(x = mtcars, from_col = 0))
  expect_equal(wb_dims(x = mtcars, from_col = 2), "B2:L33")

  # use 1 column name works

  expect_error(wb_dims(cols = "hp"), "Supplying a single argument")
  expect_error(
    wb_dims(x = mtcars, cols = c("hp", "vs")),
    regexp = "Supplying multiple column names is not supported"
  )
  expect_error(wb_dims(x = mtcars, rows = "hp"), "[Uu]se `cols` instead.")
  # Access only row / col name
  # dims of the column names of an object
  expect_equal(wb_dims(x = mtcars, rows = 0, col_names = TRUE), "A1:K1")
  expect_no_message(wb_dims(x = mtcars, rows = 0))
  # to write without column names, specify `from_row = 0` (or -1 of what you wanted)
})

test_that("`wb_dims()` handles row_names = TRUE consistenly.", {
  skip("selecting only row_names is not well supported")
  # Works well for selecting data
  expect_equal(wb_dims(x = mtcars, row_names = TRUE), "B2:L33")

  expect_equal(
    wb_dims(x = mtcars, row_names = TRUE, col_names = FALSE, cols = 0),
    wb_dims(x = mtcars, row_names = TRUE, cols = 0)
  )
  expect_equal(wb_dims(x = mtcars, row_names = TRUE), "A1:L33")
  expect_equal(wb_dims(x = mtcars, row_names = TRUE), "A1:L33")
  expect_error(wb_dims(x = mtcars, row_names = TRUE, col_names = FALSE, from_col = 0), "A2:L33")

  expect_equal(wb_dims(x = mtcars, row_names = TRUE, col_names = FALSE, from_row = 1), "A2:L33")
  wb_dims(x = mtcars, row_names = TRUE, col_names = FALSE, from_row = 2)

  expect_equal(out, "A1:L32")
  expect_equal(out2, "A1:L32")
  # Style row names of an object
  expect_equal(wb_dims(x = mtcars, cols = 0, row_names = TRUE), "A1:L33")
  expect_equal(wb_dims(x = mtcars, col_names = FALSE, row_names = TRUE), "B2:L33")
  expect_equal(wb_dims(x = mtcars, col_names = TRUE, row_names = TRUE), "B2:L34")
  # to write without column names on top
  expect_equal(wb_dims(x = mtcars, col_names = FALSE, row_names = TRUE, from_row = 0), "A1:L33")
  # to select data with row names
  expect_equal(wb_dims(x = mtcars, col_names = FALSE, row_names = TRUE), "B1:L33")
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
  expect_equal("HmPsw2WtYSxSgZ6t", tmp)

  x <- .Random.seed
  tmp <- random_string(length = 6)
  y <- .Random.seed
  expect_identical(x, y)
  expect_equal("GNZuCt", tmp)

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
  expect_equal(exp, got)

  exp <- "rId2"
  got <- get_relship_id(obj = wb$worksheets_rels[[1]], x = "drawing")
  expect_equal(exp, got)

})

test_that("as_xml_attr works", {

  mm <- matrix("", 2, 2)
  mm[1, 1] <- openxlsx2:::as_xml_attr(TRUE)
  mm[2, 1] <- openxlsx2:::as_xml_attr(FALSE)

  mm[1, 2] <- openxlsx2:::as_xml_attr(1)
  mm[2, 2] <- openxlsx2:::as_xml_attr(NULL)

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

})

test_that("outdec = \",\" works", {

  options(OutDec = ",")

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
  exp <- "[1] \"Hello World \""
  expect_equal(exp, got)

  ## watch out!
  txt <- fmt_txt("Sum ", bold = TRUE) + 2 + 2
  exp <- "Sum 22"
  got <- as.character(txt)
  expect_equal(exp, got)

  txt <- fmt_txt("Sum ", bold = TRUE) + (2 + 2)
  exp <- "Sum 4"
  got <- as.character(txt)
  expect_equal(exp, got)

})
