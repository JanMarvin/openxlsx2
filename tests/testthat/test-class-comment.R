testsetup()

test_that("class wbComment works", {
  expect_null(assert_comment(wb_comment()))
})

test_that("wb_comment and create_comment are the same except for the different defaults", {
  c1 <- create_comment("x1", author = "")
  c1_wb <- wb_comment("x1", visible = TRUE, author = "")
  expect_equal(c1, c1_wb)
  # create_comment drops multiple widths and heights silently.
  # wb_comment errors in this case
  expect_silent(create_comment(text = "x", author = "", width = c(1, 2)))
  expect_error(wb_comment(text = "x", author = "", width = c(1, 2)), "width must be a single")

})

test_that("create_comment() works", {
  # error checking
  expect_silent(create_comment("hi", width = 1))
  expect_silent(create_comment("hi", width = 1L))
  expect_silent(create_comment("hi", width = c(1, 2)))
  expect_silent(create_comment("hi", width = 1:2))
  expect_error(wb_comment("hi", width = 1:2), regexp = "width must be a")

  expect_silent(create_comment("hi", height = 1))
  expect_silent(create_comment("hi", height = 1L))
  expect_silent(create_comment("hi", height = 1:2))
  expect_error(wb_comment("hi", height = 1:2))
  expect_error(create_comment("hi", visible = NULL))
  expect_error(create_comment("hi", visible = c(TRUE, FALSE)))

  expect_error(create_comment("hi", author = 1))
  expect_error(create_comment("hi", author = c("a", "a")))

  expect_s3_class(create_comment("Hello"), "wbComment")
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

  # write on second sheet
  tmp <- temp_xlsx()
  wb <- wb_workbook()
  wb$add_worksheet()
  wb$add_worksheet()

  # write comment without author
  c1 <- create_comment(text = "this is a comment", author = "", visible = FALSE)
  wb$add_comment(dims = "B10", comment = c1)

  expect_silent(wb$save(tmp))
})


test_that("load comments", {

  fl <- testfile_path("pivot_notes.xlsx")
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
    wb_add_comment(dims = "A1", comment = c1)

  expect_equal(wb$comments, wb2$comments)


  expect_error(
    wb_workbook()$add_worksheet()$add_comment(dims = "A1"),
    'argument "comment" is missing, with no default'
  )
})

test_that("wb_add_comment() works without supplying a wbComment object.", {
  # Do not alter the workspace

  op <- options("openxlsx2.creator" = "user")
  on.exit(options(op), add = TRUE)

  # Using the new default values of wb_comment() (options("openxlsx2.creators))
  wb <- wb_workbook()$add_worksheet()$add_comment(comment = "this is a comment", dims = "A1")

  c2 <- wb_comment(text = "this is a comment")
  wb2 <- wb_workbook() %>%
    wb_add_worksheet() %>%
    wb_add_comment(dims = "A1", comment =  c2)
  # wb_comment() defaults and comment = "text" defaults are the same.
  expect_equal(wb$comments, wb2$comments)

  # The wrapper behaves the same
  wb3 <- wb_workbook()$add_worksheet()$add_comment(comment = "this is a comment")
  expect_equal(wb$comments, wb3$comments)

})


test_that("wb_remove_comment", {

  c1 <- create_comment(text = "this is a comment", author = "")
  wb <- wb_workbook()$
    add_worksheet()$
    add_comment(dims = "A1", comment = c1)$
    remove_comment(dims = "A1")

  # deprecated col / row code
  wb2 <- wb_workbook() %>% wb_add_worksheet()
  expect_warning(
    wb2 <- wb2 %>%
      wb_add_comment(col = "A", row = 1, comment = c1),
    "'col/row' is deprecated."
  )
  expect_warning(
    wb2 <- wb2 %>% wb_remove_comment(col = "A", row = 1),
    "'col/row/gridExpand' is deprecated."
  )

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
    add_comment(dims = "B10", comment = c1)$
    add_worksheet()$
    remove_worksheet(1)

  expect_silent(wb$save(temp))

})

test_that("fmt_txt in comment", {

  tmp <- temp_xlsx()
  txt <- fmt_txt("Hello ", bold = TRUE) + fmt_txt("World")
  c1 <- create_comment(text = txt, author = "bla")

  wb <- wb_workbook()$add_worksheet()$add_comment(dims = "B10", comment = c1)
  expect_silent(wb$save(tmp))

  wb <- wb$load(tmp)

  exp <- c(
    "<t xml:space=\"preserve\">bla:</t>", "<t xml:space=\"preserve\">\n</t>",
    "<t xml:space=\"preserve\">Hello </t>", "<t>World</t>"
  )
  got <- wb$comments[[1]][[1]]$comment
  expect_equal(exp, got)

})

test_that("threaded comments work", {

  wb <- wb_workbook()$add_worksheet()

  wb$add_person(name = "Kirk")
  wb$add_person(name = "Uhura")
  wb$add_person(name = "Spock")
  wb$add_person(name = "Scotty")

  kirk_id <- wb$get_person(name = "Kirk")$id
  uhura_id <- wb$get_person(name = "Uhura")$id
  spock_id <- wb$get_person(name = "Spock")$id
  scotty_id <- wb$get_person(name = "Scotty")$id

  # write a comment to a thread, reply to one and solve some
  wb <- wb %>%
    wb_add_thread(dims = "A1", comment = "wow it works!", person_id = kirk_id) %>%
    wb_add_thread(dims = "A2", comment = "indeed", person_id = uhura_id, resolve = TRUE) %>%
    wb_add_thread(dims = "A1", comment = "fascinating", person_id = spock_id, reply = TRUE)

  exp <- data.frame(
    ref = c("A1", "A1"),
    displayName = c("Kirk", "Spock"),
    text = c("wow it works!", "fascinating"),
    done = c("0", "")
  )
  got <- wb_get_thread(wb)[, -1]
  # somehow the row ordering differs for parallel and non-parallel testthat runs
  expect_equal(exp[order(got$displayName), ], got, ignore_attr = TRUE)

  exp <- "[Threaded comment]\n\nYour spreadsheet software allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of a certain spreadsheet software.\n\nComment: wow it works!\nReplie:fascinating"
  got <- wb_get_comment(wb)$comment
  expect_equal(exp, got)

  # start a new thread
  wb <- wb %>%
    wb_add_thread(dims = "A1", comment = "oops", person_id = kirk_id)

  exp <- data.frame(
    ref = "A1",
    displayName = "Kirk",
    text = "oops",
    done = "0"
  )
  got <- wb_get_thread(wb)[, -1]
  expect_equal(exp, got)

  wb <- wb %>%
    wb_add_worksheet() %>%
    wb_add_thread(dims = "A1", comment = "hmpf", person_id = scotty_id)

  exp <- data.frame(
    ref = "A1",
    displayName = "Scotty",
    text = "hmpf",
    done = "0"
  )
  got <- wb_get_thread(wb)[, -1]
  expect_equal(exp, got)

})

test_that("thread option works", {

  wb <- wb_workbook()$add_worksheet()
  wb$add_person(name = "Kirk")
  wb <- wb %>% wb_add_thread(comment = "works")

  exp <- "works"
  got <- wb_get_thread(wb)$text
  expect_equal(exp, got)

})
