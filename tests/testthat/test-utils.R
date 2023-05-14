
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

  wb <- wb_load(system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))

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
