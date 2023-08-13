test_that("fill merged cells", {
  wb <- wb_workbook()
  wb$add_worksheet("sheet1")
  wb$add_data(1, data.frame("A" = 1, "B" = 2))
  wb$add_data(1, 2, startRow = 2, startCol = 2)
  wb$add_data(1, 3, startRow = 2, startCol = 3)
  wb$add_data(1, 4, startRow = 2, startCol = 4)
  wb$add_data(1, t(matrix(1:4, 4, 4)), startRow = 3, startCol = 1, colNames = FALSE)

  wb$merge_cells(1, rows = 1, cols = 2:4)
  wb$merge_cells(1, rows = 3, cols = 2:4)
  wb$merge_cells(1, rows = 4, cols = 2:4)
  wb$merge_cells(1, rows = 5, cols = 2:4)

  tmp_file <- temp_xlsx()
  wb_save(wb, tmp_file)

  # in openxlsx X3 and X4 because of name fixing
  expect_equal(names(read_xlsx(tmp_file, fillMergedCells = FALSE)), c("A", "B", NA_character_, NA_character_))
  expect_equal(names(read_xlsx(tmp_file, fillMergedCells = TRUE)), c("A", "B", "B", "B"))

  r1 <- data.frame("A" = rep(1, 5), "B" = rep(2, 5), "NA1" = rep(3, 5), "NA2" = rep(4, 5))
  rnams <- as.character(seq(2, length.out = nrow(r1)))
  dimnames(r1) <-  list(rnams, c("A", "B", NA_character_, NA_character_))

  r2 <- data.frame("A" = rep(1, 5), "B" = rep(2, 5), "B1" = c(3, 2, 2, 2, 3), "B2" = c(4, 2, 2, 2, 4))
  dimnames(r2) <-  list(rnams, c("A", "B", "B", "B"))

  r2_1 <- r2[1:5, 1:3]
  names(r2_1) <- c("A", "B", "B")

  expect_equal(read_xlsx(tmp_file, fillMergedCells = FALSE), r1, ignore_attr = TRUE)
  expect_equal(read_xlsx(tmp_file, fillMergedCells = TRUE), r2, ignore_attr = TRUE)

  expect_equal(read_xlsx(tmp_file, cols = 1:3, fillMergedCells = TRUE), r2_1, ignore_attr = TRUE)
  expect_equal(read_xlsx(tmp_file, rows = 1:3, fillMergedCells = TRUE), r2[1:2, ], ignore_attr = TRUE)
  expect_equal(read_xlsx(tmp_file, cols = 1:3, rows = 1:4, fillMergedCells = TRUE), r2_1[1:3, ], ignore_attr = TRUE)
})

test_that("merge and unmerge cells", {

  wb <- wb_workbook()$add_worksheet()$merge_cells(rows = 1:2, cols = 1:2)

  expect_error(wb$merge_cells(rows = 1:2, cols = 1:2), "Remove existing merge first.")
  expect_silent(wb$unmerge_cells(rows = 1:2, cols = 1:2))
  expect_silent(wb$merge_cells(rows = 1:2, cols = 1:2))

})

test_that("fill merged NA cells", {
  wb <- wb_workbook()
  wb$add_worksheet("sheet1")
  wb$add_data(1, t(matrix(c(1:3, NA_real_), 4, 4)), startRow = 3, startCol = 1, colNames = FALSE)

  wb$merge_cells(1, rows = 1:4, cols = 4)

  tmp_file <- temp_xlsx()
  wb_save(wb, tmp_file)

  r1 <- t(matrix(c(1:3, NA_real_), 4, 4))
  expect_equal(as.matrix(read_xlsx(tmp_file, fillMergedCells = FALSE,
                                   rowNames = FALSE, colNames = FALSE)),
               r1,
               ignore_attr = TRUE)

  expect_equal(as.matrix(read_xlsx(tmp_file, fillMergedCells = TRUE,
                                   rowNames = FALSE, colNames = FALSE)),
               r1,
               ignore_attr = TRUE)
})

test_that("solving merge conflicts works", {
  wb <- wb_workbook()$add_worksheet()
  wb$merge_cells(dims = "A1:B2")
  wb$merge_cells(dims = "G1:K2")

  # first solve replacing A1:B2
  wb$merge_cells(dims = "A1:C3", solve = TRUE)

  exp <- c("<mergeCell ref=\"G1:K2\"/>", "<mergeCell ref=\"A1:C3\"/>")
  got <- wb$worksheets[[1]]$mergeCells
  expect_equal(exp, got)

  # second solve replace A1:C3 and parts of G1:K2
  wb$merge_cells(dims = "A1:J7", solve = TRUE)

  exp <- c("<mergeCell ref=\"K1:K2\"/>", "<mergeCell ref=\"A1:J7\"/>")
  got <- wb$worksheets[[1]]$mergeCells
  expect_equal(exp, got)

  # third solve replace both parts with single merge
  wb$merge_cells(dims = "A1:Z8", solve = TRUE)

  exp <- "<mergeCell ref=\"A1:Z8\"/>"
  got <- wb$worksheets[[1]]$mergeCells
  expect_equal(exp, got)

  # fourth solve insert into merged cell region
  wb$merge_cells(dims = "B2:D4", solve = TRUE)

  exp <- c(
    "<mergeCell ref=\"A1:A8\"/>", "<mergeCell ref=\"E1:Z8\"/>",
    "<mergeCell ref=\"B1:D1\"/>", "<mergeCell ref=\"B5:D8\"/>",
    "<mergeCell ref=\"B2:D4\"/>"
  )
  got <- wb$worksheets[[1]]$mergeCells
  expect_equal(exp, got)

})
