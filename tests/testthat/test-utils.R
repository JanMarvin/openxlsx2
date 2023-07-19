
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
  expect_equal(wb_dims(1L, 1L), "A1")
  expect_equal(wb_dims(1:10, 1:26), "A1:Z10")
  expect_equal(wb_dims(1:10, LETTERS), "A1:Z10")
  expect_equal(
    wb_dims(1:10, 1:12, start_row = 2),
    wb_dims(rows = 1:10, cols = 1:12, start_row = 2)
  )

  # Ambiguous / input not accepted.
  # This now fails, as it used not to work. (Use `wb_dims()`, `NULL`, or )
  expect_error(wb_dims(NULL), "Specifying a single")
  expect_error(wb_dims(1), "Specifying a single unnamed argument is not handled")
  # This used to return A1 as well.
  expect_error(wb_dims(2), "Specifying a single unnamed argument is not handled")
  expect_error(wb_dims(mtcars), "Specifying a single unnamed argument")

  skip_on_ci("`wb_dims()` WIP")
  expect_error(wb_dims(rows = c(1, 3, 4), cols = c(1, 4)), "wb_dims() should only be used for specifying a single continuous range.")
})


test_that("wb_dims() works when not specifying an object.", {
  expect_equal(wb_dims(rows =  1:10, cols = 5:7), "E1:G10")
  expect_equal(wb_dims(rows = 5:7, cols =  1:10), "A5:J7")
  expect_equal(wb_dims(rows = 5, cols = 7), "G5")
  expect_error(
    wb_dims(cols =  1:10, col = 5:7),
    "found only one cols/rows argument"
  )
  expect_equal(wb_dims(1:2, 1:4, start_row = 2, start_col = "B"), "B2:E3")

  skip_on_ci("`wb_dims()` WIP")
  expect_equal(wb_dims(start_col = 4), "D1")
  expect_equal(wb_dims(start_row = 4), "A4")
  expect_equal(wb_dims(start_row = 4, start_col = 3), "C4")
  expect_equal(wb_dims(4, 3), wb_dims(start_row = 4, start_col = 3))

})

test_that("`wb_dims()` works when specifying an object `x`.",{
  expect_equal(wb_dims(x = mtcars), "A1:K33")
  expect_equal(wb_dims(x = mtcars, col_names = FALSE, row_names = TRUE), "A1:L32")

  expect_equal(wb_dims(x = letters), "A1:A26")

  expect_equal(wb_dims(x = t(letters)), "A1:Z2")

  expect_equal(wb_dims(x = mtcars, start_row = 2, start_col = "B"), "B2:L34")

  skip_on_ci("`wb_dims()` WIP")
  # use `col_names = FALSE` as a way to access the data, when formatting content only
  # currently
  # wb_dims(x = mtcars, col_names = FALSE) = "A1:K32"
  # proposed. TODO
  expect_equal(wb_dims(x = mtcars, col_names= FALSE), "A2:K33")
  expect_equal(wb_dims(rows = 1:(nrow(mtcars) + 1), cols = 4), "D1:D33")
  expect_equal(wb_dims(x = mtcars, cols = 4), "D1:D33")
  expect_equal(wb_dims(x = mtcars, cols = 4),  "D1:D33")

  expect_equal(wb_dims(x = mtcars, col_names = FALSE, start_col = 2), "B2:L33")
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
