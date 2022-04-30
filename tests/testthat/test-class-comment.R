
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
  s1 <- create_font(b = "true", color = c(rgb = "FFFF0000"), sz = "12")
  s2 <- create_font(color = c(rgb = "FF000000"), sz = "9")
  c3 <- create_comment(text = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))

  expect_silent(write_comment(wb, 1, col = 6, row = 3, comment = c3))

  expect_true(length(wb$comments) == 1)
  expect_true(length(wb$comments[[1]]) == 3)

  expect_silent(remove_comment(wb, 1, col = "B", row = 10))

  expect_true(length(wb$comments) == 1)
  expect_true(length(wb$comments[[1]]) == 2)

  expect_silent(wb_save(wb, tmp))

})
