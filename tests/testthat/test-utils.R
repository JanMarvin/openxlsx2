
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
