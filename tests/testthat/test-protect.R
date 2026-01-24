test_that("Protect Workbook", {

  op <- options("openxlsx2.legacy_password" = TRUE)
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()
  wb$add_worksheet("s1")

  wb$protect(password = "abcdefghij")

  expect_equal(wb$workbook$workbookProtection, "<workbookProtection lockStructure=\"0\" lockWindows=\"0\" hashPassword=\"FEF1\"/>")

  wb$protect(password = "abcdefghij", file_sharing = TRUE)
  expect_true(
    grepl("<fileSharing userName=\".*?\" reservationPassword=\"FEF1\"/>",
          wb$workbook$fileSharing)
  )

  # this creates a corrupted workbook:
  # 1) this nulls the reference
  wb$protect(protect = FALSE, password = "abcdefghij", lock_structure = TRUE, lock_windows = TRUE)
  expect_equal(wb$workbook$workbookProtection, NULL)

  # 2) this creates the reference, but at the wrong position (at the end, not at 6)
  wb$protect(password = "abcdefghij", lock_structure = TRUE)

  wb$protect(protect = TRUE, password = NULL)
  expect_equal(wb$workbook$workbookProtection, "<workbookProtection lockStructure=\"0\" lockWindows=\"0\"/>")

  expect_warning(
    wb$protect(password = paste0(letters, collapse = ""), file_sharing = TRUE),
    "Excel password protection only uses the first 15 characters."
  )
})

test_that("sha512 hashing works", {

  skip_if_not_installed("openssl")

  wb <- wb_workbook()
  wb$add_worksheet("s1")
  wb$protect(password = "abcdefghij", file_sharing = TRUE)

  exp <- c("lockStructure", "lockWindows", "algorithmName", "hashPassword",
           "saltValue", "spinCount")
  got <- names(xml_attr(wb$workbook$workbookProtection, "workbookProtection")[[1]])
  expect_equal(got, exp)

  exp <- c("userName", "algorithmName", "spinCount", "hashValue", "saltValue")
  got <- names(xml_attr(wb$workbook$fileSharing, "fileSharing")[[1]])
  expect_equal(got, exp)

  wb <- wb_protect_worksheet(wb, sheet = "s1", protect = TRUE, password = "abcdefghij", properties =  c("formatCells", "formatColumns", "PivotTables"))

  exp <- c("sheet", "formatCells", "formatColumns", "algorithmName", "spinCount",
           "hashValue", "saltValue")
  got <- names(xml_attr(wb$worksheets[[1]]$sheetProtection, "sheetProtection")[[1]])
  expect_equal(got, exp)

})

test_that("Reading protected Workbook", {
  temp_file <- temp_xlsx()
  on.exit(unlink(temp_file), add = TRUE)

  wb <- wb_workbook()
  wb$add_worksheet("s1")
  wb_protect(wb, password = "abcdefghij")
  wb_save(wb, temp_file)

  wb2 <- wb_load(file = temp_file)
  # Check that the order of the sub-elements is preserved
  n1 <- names(wb2$workbook)
  n2 <- names(wb$workbook)[names(wb$workbook) != "apps"]
  expect_equal(n1, n2)

})

test_that("Protection", {

  op <- options("openxlsx2.legacy_password" = TRUE)
  on.exit(options(op), add = TRUE)

  wb <- wb_workbook()
  wb$add_worksheet("s1")
  wb$add_worksheet("s2")

  wb <- wb_protect_worksheet(wb, sheet = "s1", protect = TRUE, password = "abcdefghij", properties =  c("formatCells", "formatColumns", "PivotTables"))
  expect_true(wb$worksheets[[1]]$sheetProtection == "<sheetProtection sheet=\"1\" formatCells=\"1\" formatColumns=\"1\" password=\"FEF1\"/>")

  wb <- wb_protect_worksheet(wb, sheet = "s2", protect = TRUE)
  expect_true(wb$worksheets[[2]]$sheetProtection == "<sheetProtection sheet=\"1\"/>")

  wb <- wb_protect_worksheet(wb, sheet = "s2", protect = FALSE)
  expect_equal(wb$worksheets[[2]]$sheetProtection, character())
})

test_that("logical vector works too", {

  wb <- wb_workbook()$add_worksheet("S1")

  # pre 0.7.1 style
  wb$protect_worksheet(
    "S1",
    protect = TRUE,
    properties = c("formatCells", "formatColumns", "insertColumns", "deleteColumns")
  )

  exp <- "<sheetProtection sheet=\"1\" formatCells=\"1\" formatColumns=\"1\" insertColumns=\"1\" deleteColumns=\"1\"/>"
  got <- wb$worksheets[[1]]$sheetProtection
  expect_equal(got, exp)

  # post 0.7 style
  wb$protect_worksheet(
    "S1",
    protect = TRUE,
    properties = c(formatCells = FALSE, formatColumns = FALSE, insertColumns = TRUE, deleteColumns = TRUE)
  )

  exp <- "<sheetProtection sheet=\"1\" formatCells=\"0\" formatColumns=\"0\" insertColumns=\"1\" deleteColumns=\"1\"/>"
  got <- wb$worksheets[[1]]$sheetProtection
  expect_equal(got, exp)

})
