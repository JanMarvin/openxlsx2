testsetup()

test_that("class wbComment works", {
  expect_null(assert_comment(wb_comment()))
})

test_that("wb_comment() works", {
  # error checking
  expect_error(wb_comment("hi", width = 1:2), regexp = "width must be a")

  expect_error(wb_comment("hi", height = 1:2))

  expect_error(wb_comment(text = "x", author = "", width = c(1, 2)), "width must be a single")
})


test_that("comments", {
  tmp <- temp_xlsx()
  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")

  # write comment without author
  c1 <- wb_comment(text = "this is a comment", author = "")
  wb <- wb_add_comment(wb, 1, dims = "B10", comment = c1)

  # Write another comment with author information
  c2 <- wb_comment(text = "this is another comment", author = "Marco Polo")
  wb <- wb_add_comment(wb, 1, dims = "C10", comment = c2)

  # write a styled comment with system author
  s1 <- create_font(b = "true", color = wb_colour(hex = "FFFF0000"), sz = "12")
  s2 <- create_font(color = wb_colour(hex = "FF000000"), sz = "9")
  c3 <- wb_comment(text = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))

  expect_silent(wb$add_comment(1, dims = "F1", comment = c3))

  expect_length(wb$comments, 1)
  expect_length(wb$comments[[1]], 3)

  expect_silent(wb$remove_comment(1, dims = "B10"))

  expect_length(wb$comments, 1)
  expect_length(wb$comments[[1]], 2)

  expect_silent(wb_save(wb, tmp))

  # write on second sheet
  tmp <- temp_xlsx()
  wb <- wb_workbook()
  wb$add_worksheet()
  wb$add_worksheet()

  # write comment without author
  c1 <- wb_comment(text = "this is a comment", author = "")
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
  c1 <- wb_comment(text = "this is a comment", author = "")
  expect_silent(wb$add_comment(5, dims = "B10", comment = c1))

  wb$save(temp)

  tempd <- temp_dir("comment_ext")
  unzip(temp, exdir = tempd)
  comments <- dir(path = paste0(tempd, "/xl"), pattern = "comment")

  expect_equal(c("comments1.xml", "comments2.xml", "comments3.xml"), comments)
  unlink(tempd, recursive = TRUE)

})

test_that("wb_add_comment", {

  c1 <- wb_comment(text = "this is a comment", author = "")

  wb <- wb_workbook()$add_worksheet()$add_comment(dims = "A1", comment = c1)

  wb2 <- wb_workbook()$
    add_worksheet()$
    add_comment(dims = "A1", comment = c1)

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
  wb2 <- wb_workbook()$
    add_worksheet()$
    add_comment(dims = "A1", comment =  c2)
  # wb_comment() defaults and comment = "text" defaults are the same.
  expect_equal(wb$comments, wb2$comments)

  # The wrapper behaves the same
  wb3 <- wb_workbook()$add_worksheet()$add_comment(comment = "this is a comment")
  expect_equal(wb$comments, wb3$comments)

})


test_that("wb_remove_comment", {

  c1 <- wb_comment(text = "this is a comment", author = "")
  wb <- wb_workbook()$
    add_worksheet()$
    add_comment(dims = "A1", comment = c1)$
    remove_comment(dims = "A1")

  # deprecated col / row code
  wb2 <- wb_workbook()$add_worksheet()
  expect_warning(
    wb2$add_comment(col = "A", row = 1, comment = c1),
    "'col/row' is deprecated."
  )
  expect_warning(
    wb2$remove_comment(col = "A", row = 1),
    "'col/row/gridExpand' is deprecated."
  )

  expect_equal(wb$comments, wb2$comments)

})

test_that("print comment", {

  c2 <- wb_comment(text = "this is another comment",
                   author = "Marco Polo")
  got <- capture_output(print(c2), print = TRUE)
  exp <- "Author: Marco Polo\nText:\n Marco Polo:\nthis is another comment\n\nStyle:\n\n\n\n\nFont name: Aptos Narrow\nFont size: 11\nFont color: #000000\n\n"
  expect_equal(got, exp)

})

test_that("removing comment sheet works", {

  temp <- temp_xlsx()
  c1 <- wb_comment(text = "this is a comment", author = "")

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
  c1 <- wb_comment(text = txt, author = "bla")

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
  wb$
    add_thread(dims = "A1", comment = "wow it works!", person_id = kirk_id)$
    add_thread(dims = "A2", comment = "indeed", person_id = uhura_id, resolve = TRUE)$
    add_thread(dims = "A1", comment = "fascinating", person_id = spock_id, reply = TRUE)

  exp <- data.frame(
    ref = c("A1", "A1"),
    displayName = c("Kirk", "Spock"),
    text = c("wow it works!", "fascinating"),
    done = c("0", ""),
    stringsAsFactors = FALSE
  )
  got <- wb_get_thread(wb, dims = "A1")[, -1]
  # somehow the row ordering differs for parallel and non-parallel testthat runs
  expect_equal(exp[order(got$displayName), ], got, ignore_attr = TRUE)

  exp <- "[Threaded comment]\n\nYour spreadsheet software allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of a certain spreadsheet software.\n\nComment: wow it works!\nReplie:fascinating"
  got <- wb_get_comment(wb, dims = "A1")$comment
  expect_equal(exp, got)

  # start a new thread
  wb$add_thread(dims = "A1", comment = "oops", person_id = kirk_id)

  exp <- data.frame(
    ref = "A1",
    displayName = "Kirk",
    text = "oops",
    done = "0",
    stringsAsFactors = FALSE
  )
  got <- wb_get_thread(wb, dims = "A1")[, -1]
  expect_equal(exp, got)

  wb$add_worksheet()$
    add_thread(dims = "A1", comment = "hmpf", person_id = scotty_id)

  exp <- data.frame(
    ref = "A1",
    displayName = "Scotty",
    text = "hmpf",
    done = "0",
    stringsAsFactors = FALSE
  )
  got <- wb_get_thread(wb, dims = "A1")[, -1]
  expect_equal(exp, got)

})

test_that("thread option works", {

  wb <- wb_workbook()$add_worksheet()
  wb$add_person(name = "Kirk")
  wb$add_thread(comment = "works")

  exp <- "works"
  got <- wb_get_thread(wb)$text
  expect_equal(exp, got)

})

test_that("background images work", {

  wb <- wb_workbook()
  wb$add_worksheet("Sheet 1")

  # file extension must be png or jpeg, not jpg?
  tmp <- tempfile(fileext = ".png")
  png(filename = tmp, bg = "transparent")
  plot(1:10)
  rect(1, 5, 3, 7, col = "white")
  dev.off()

  # write comment without author
  c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
  wb$add_comment(dims = "B12", comment = c1, file = tmp)

  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  wb$add_worksheet()$add_image(dims = "C5", file = img, width = 6, height = 5)
  wb$add_comment(dims = "B12", comment = c1)

  wb$add_worksheet()

  # file extension must be png or jpeg, not jpg?
  tmp2 <- tempfile(fileext = ".png")
  png(filename = tmp2, bg = "transparent")
  barplot(1:10)
  dev.off()

  # write comment without author
  c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
  wb$add_comment(dims = "G12", comment = c1, file = tmp2)
  wb$add_comment(dims = "G12", sheet = 1, comment = c1, file = tmp2)

  expect_length(wb$vml, 3)
  expect_length(wb$vml_rels, 3)
  expect_length(wb$vml_rels[[1]], 2)
  expect_null(wb$vml_rels[[2]])
  expect_length(wb$vml_rels[[3]], 1)

})

test_that("More than two background images work", {

  tmp <- tempfile(fileext = ".png")
  png(filename = tmp, bg = "transparent")
  plot(1:10)
  dev.off()

  c1 <- wb_comment(text = "Comm1", author = "", visible = TRUE)

  wb <- wb_workbook()$
    add_worksheet()$
    add_comment(dims = "A2", comment = c1, file = tmp)$
    add_comment(dims = "A3", comment = c1, file = tmp)$
    add_comment(dims = "A4", comment = c1, file = tmp)$
    add_worksheet()$
    add_comment(dims = "A2", comment = c1, file = tmp)

  exp <- list(
    c(
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/>",
      "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image2.png\"/>",
      "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image3.png\"/>"
    ),
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image4.png\"/>"
  )
  got <- wb$vml_rels
  expect_equal(exp, got)

})

test_that("background colors work", {

  wb <- wb_workbook()$add_worksheet()

  txt <- fmt_txt("This Part Bold red\n\n", bold = TRUE, size = 12, color = wb_color("red")) +
    fmt_txt("This part black", size = 9, color = wb_color("black"))

  wb$add_comment(sheet = 1, dims = wb_dims(3, 6), comment = wb_comment(text = txt), color = wb_color("green"))

  expect_match(wb$vml[[1]], "fillcolor=\"#00FF00\"")

})

test_that("wb_get_comment works", {
  wb <- wb_workbook()$
    add_worksheet()

  # Add a comment to cell A1
  c1 <- wb_comment(text = "This is a sample comment.", author = "openxlsx2 authors")

  wb$add_comment(
    dims = "B2",
    comment = c1
  )

  exp <- data.frame(ref = "B2",
                    author = "openxlsx2 authors",
                    comment = "openxlsx2 authors: \n This is a sample comment.",
                    cmmt_id = 1L,
                    stringsAsFactors = FALSE)
  got <- wb$get_comment(dims = "B2")
  expect_equal(exp, got)

  wb$remove_comment(dims = "B2")
  expect_equal(list(), wb$get_comment(dims = "B2"))

  wb$add_comment(
    dims = "B2",
    comment = "foo"
  )
  got <- wb$get_comment(dims = "B2")
  expect_equal(1L, nrow(got))
})

test_that("removing comment from second sheet works", {

  temp <- temp_xlsx()

  c1 <- wb_comment(text = "this is a comment", author = "")

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_worksheet("Sheet 2")$
    add_comment(dims = "B10", comment = c1)$
    remove_comment(dims = "B10")

  expect_silent(wb$save(temp))

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_worksheet("Sheet 2")$
    add_comment(dims = "B10", comment = c1)$
    add_comment(dims = "D10", comment = c1)$
    add_comment(dims = "B12", comment = c1)

  wb$remove_comment(sheet = "Sheet 2", dims = "B10")
  wb$remove_comment(sheet = "Sheet 2", dims = "D10")
  expect_warning(wb$remove_comment(sheet = "Sheet 2", dims = "G12"), "no comment to remove")
  wb$remove_comment(sheet = "Sheet 2", dims = "B12")

  expect_silent(wb$save(temp))
  expect_equal(list(), wb$comments)
  expect_equal(list(), wb$vml)

})
