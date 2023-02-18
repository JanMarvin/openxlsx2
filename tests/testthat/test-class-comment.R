
test_that("class wbComment works", {
  expect_null(assert_comment(wb_comment()))
})

test_that("create_comment() works", {
  # error checking
  expect_silent(create_comment("hi", width = 1))
  expect_error(create_comment("hi", width = 1L))
  expect_error(create_comment("hi", width = 1:2))

  expect_silent(create_comment("hi", height = 1))
  expect_error(create_comment("hi", height = 1L))
  expect_error(create_comment("hi", height = 1:2))

  expect_error(create_comment("hi", visible = NULL))
  expect_error(create_comment("hi", visible = c(TRUE, FALSE)))

  expect_error(create_comment("hi", author = 1))
  expect_error(create_comment("hi", author = c("a", "a")))

  expect_true(inherits(create_comment("Hello"), "wbComment"))
})


test_that("comments", {
  tmp <- temp_xlsx()
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")

  # write comment without author
  c1 <- create_comment(text = "this is a comment", author = "")
  write_comment(wb, 1, col = "B", row = 10, comment = c1)

  # Write another comment with author information
  c2 <- create_comment(text = "this is another comment", author = "Marco Polo")
  write_comment(wb, 1, col = "C", row = 10, comment = c2)

  # write a styled comment with system author
  s1 <- create_font(b = "true", color = wb_colour(hex = "FFFF0000"), sz = "12")
  s2 <- create_font(color = wb_colour(hex = "FF000000"), sz = "9")
  c3 <- create_comment(text = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))

  expect_silent(write_comment(wb, 1, col = 6, row = 3, comment = c3))

  expect_true(length(wb$comments) == 1)
  expect_true(length(wb$comments[[1]]) == 3)

  expect_silent(remove_comment(wb, 1, col = "B", row = 10))

  expect_true(length(wb$comments) == 1)
  expect_true(length(wb$comments[[1]]) == 2)

  expect_silent(wb_save(wb, tmp))

})


test_that("load comments", {

  fl <- system.file("extdata", "pivot_notes.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)

  temp <- temp_xlsx()
  wb$save(temp)

  tempd <- temp_dir("comment_ext")
  unzip(temp, exdir = tempd)
  comments <- dir(path = paste0(tempd, "/xl"), pattern = "comment")

  expect_equal(c("comments1.xml", "comments2.xml"), comments)
  unlink(tempd, recursive = TRUE)

  ## add a new comment to a workbook that has comments
  c1 <- create_comment(text = "this is a comment", author = "")
  expect_silent(write_comment(wb, 5, col = "B", row = 10, comment = c1))

  wb$save(temp)

  tempd <- temp_dir("comment_ext")
  unzip(temp, exdir = tempd)
  comments <- dir(path = paste0(tempd, "/xl"), pattern = "comment")

  expect_equal(c("comments1.xml", "comments2.xml", "comments3.xml"), comments)
  unlink(tempd, recursive = TRUE)

})

test_that("wb_add_comment", {

  c1 <- create_comment(text = "this is a comment", author = "")

  wb <- wb_workbook()$add_worksheet()$add_comment(dims = "A1", comment = c1)

  wb2 <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_comment(col = "A", row = 1, comment = c1)

  expect_equal(wb$comments, wb2$comments)


  expect_error(
    wb_workbook()$add_worksheet()$add_comment(dims = "A1"),
    'argument "comment" is missing, with no default'
  )

})

test_that("wb_remove_comment", {

  c1 <- create_comment(text = "this is a comment", author = "")
  wb <- wb_workbook()$
    add_worksheet()$
    add_comment(dims = "A1", comment = c1)$
    remove_comment(dims = "A1")

  wb2 <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_comment(col = "A", row = 1, comment = c1) %>%
    wb_remove_comment(col = "A", row = 1)

  expect_equal(wb$comments, wb2$comments)

})

test_that("print comment", {

  c2 <- create_comment(text = "this is another comment",
                       author = "Marco Polo")

  exp <- "Author: Marco Polo\nText:\n Marco Polo:\nthis is another comment\n\nStyle:\n\n\n\n\nFont name: Calibri\nFont size: 11\nFont color: #000000\n\n"
  got <- capture_output(print(c2), print = TRUE)
  expect_equal(exp, got)

})

test_that("removing comment sheet works", {


  temp <- temp_xlsx()
  c1 <- create_comment(text = "this is a comment", author = "")

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_comment(1, col = "B", row = 10, comment = c1)$
    add_worksheet()$
    remove_worksheet(1)

  # # FIXME this still carries a reference to comments1.xml even though the file
  # # is no longer written. Spreadsheet software does not complain for now
  # wb$Content_Types[[10]]

  expect_silent(wb$save(temp))

})
