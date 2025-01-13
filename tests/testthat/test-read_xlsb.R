test_that("reading xlsb works", {

  skip_online_checks()

  xlsxFile <- testfile_path("openxlsx2_example.xlsb")

  wb <- wb_load(xlsxFile)
  df_xlsb <- wb_to_df(wb)

  xlsx <- system.file("extdata", "openxlsx2_example.xlsx",
                      package = "openxlsx2")
  df_xlsx <- wb_to_df(xlsx)

  expect_equal(df_xlsb, df_xlsx)

  df_xlsb <- wb_to_df(wb, show_formula = TRUE)
  df_xlsx <- wb_to_df(xlsx, show_formula = TRUE)
  expect_equal(df_xlsb, df_xlsx)

})

test_that("reading complex xlsb works", {

  skip_online_checks()

  xlsxFile <- testfile_path("hyperlink.xlsb")

  expect_message( # larger workbook
    capture_output( # unhandled conditions
      wb <- wb_load(xlsxFile)
    ),
    "importing larger workbook. please wait a moment"
  )

  # chartsheets
  exp <- c(TRUE, FALSE, FALSE, FALSE)
  got <- wb$is_chartsheet
  expect_equal(exp, got)

  # named regions
  exp <- c("Numbers", "characters", "_xlnm._FilterDatabase")
  got <- unique(wb$get_named_regions()$name)
  expect_equal(exp, got)

  # tables
  exp <- "Table1"
  got <- wb$get_tables(sheet = 2)$tab_name
  expect_equal(exp, got)

  # comments
  exp <- c(
    "<t>Jan Marvin Garbuszus:</t>",
    "<t xml:space=\"preserve\">\n</t>",
    "<t>A new note!</t>"
  )
  got <- wb$comments[[1]][[1]]$comment
  expect_equal(exp, got)

  # hyperlinks
  exp <- "Sheet1!E3"
  got <- xml_attr(wb$worksheets[[2]]$hyperlinks[[1]], "hyperlink")[[1]][["location"]]
  expect_equal(exp, got)

})

test_that("worksheets with real world formulas", {

  skip_online_checks()

  xlsxFile <- testfile_path("nhs-core-standards-for-eprr-v6.1.xlsb")

  expect_message( # larger workbook
    capture_output( # unhandled conditions
      suppressWarnings(wb <- wb_load(xlsxFile))
    ),
    "importing larger workbook. please wait a moment"
  )

  exp <- c("Control", "EPRR Core Standards", "Deep dive",
           "Interoperable capabilities ", "Lookups", "Calculations")
  got <- wb$get_sheet_names() %>% names()
  expect_equal(exp, got)

  xlsxFile <- testfile_path("array_fml.xlsb")

  wb <- wb_load(xlsxFile)

  exp <- structure(
    list(
      A = c("1", "2", NA, "6", "#VALUE!"),
      B = c("2", "3", NA, "a", NA),
      C = c(3, 4, NA, NA, NA),
      D = c(4, 5, NA, NA, NA),
      E = c(5, 6, NA, NA, NA)
    ),
    row.names = c(NA, 5L),
    class = "data.frame"
  )
  got <- wb_to_df(wb, col_names = FALSE)
  expect_equal(exp, got)

  exp <- structure(
    list(
      A = c("1", "A1+1", NA, "SUM({\"1\",\"2\",\"3\"})", "SUMIFS(A1:A2,'[1]foo'!B2:B3,{\"A\"})"),
      B = structure(c("A1+1", "B1+1", NA, "{\"a\"}", NA), class = c("character", "formula")),
      C = structure(c("B1+1", "C1+1", NA, NA, NA), class = c("character", "formula")),
      D = structure(c("C1+1", "D1+1", NA, NA, NA), class = c("character", "formula")),
      E = structure(c("D1+1", "E1+1", NA, NA, NA), class = c("character", "formula"))
    ),
    row.names = c(NA, 5L),
    class = "data.frame"
  )
  got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE)
  expect_equal(exp, got)


  xlsxFile <- testfile_path("match_escape.xlsb")

  wb <- wb_load(xlsxFile)

  exp <- structure(
    "IF(ISNA(MATCH(G681,{\"Annual\",\"\"\"5\"\"\",\"\"\"1+4\"\"\"},0)),1,MATCH(G681,{\"Annual\",\"\"\"5\"\"\",\"\"\"1+4\"\"\"},0))",
    class = c("character", "formula")
  )
  got <- wb_to_df(wb, col_names = FALSE, show_formula = TRUE)$A
  expect_equal(exp, got)

})

test_that("xlsb formulas", {

  fl <- testfile_path("formula_checks.xlsb")
  wb <- wb_load(fl)

  exp <- c("", "D2:E2", "A1,B1", "A1 A2", "1+1", "1-1", "1*1", "1/1",
          "1%", "1^1", "1=1", "1&gt;1", "1&gt;=1", "1&lt;1", "1&lt;=1",
          "1&lt;&gt;1", "+A3", "-R2", "(1)", "SUM(1, )", "1", "2.500000",
          "\"a\"", "\"A\"&amp;\"B\"", "Sheet2!B2", "'[1]Sheet3'!A2")
  got <- unique(wb$worksheets[[1]]$sheet_data$cc$f)
  expect_equal(exp, got)

  fl <- testfile_path("formula_checks.xlsx")
  xl <- wb_load(fl)

  expect_equal(
    xl$worksheets[[2]]$sheet_data$cc$f,
    wb$worksheets[[2]]$sheet_data$cc$f
  )

})

test_that("shared formulas are detected correctly", {

  xlsb <- testfile_path("formula_checks.xlsb")
  xlsx <- testfile_path("formula_checks.xlsx")

  wbb <- wb_load(xlsb)
  wbx <- wb_load(xlsx)

  expect_equal(
    wbb$worksheets[[4]]$sheet_data$cc,
    wbx$worksheets[[4]]$sheet_data$cc
  )

})

test_that("loading custom sheet view in xlsb files works", {

  skip_online_checks()

  fl <- testfile_path("custom_sheet_view.xlsb")

  wb <- wb_load(fl)

  exp <- "<customSheetViews><customSheetView guid=\"{101E93B2-5AEC-EA4B-A037-FE6837BBF571}\" view=\"pageLayout\"><selection pane=\"topLeft\" activeCell=\"E23\" sqref=\"E23\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{29911425-4959-8744-A301-8A3516CFBBC5}\" view=\"pageLayout\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{E70FF13A-F852-294B-B121-766AE66F850F}\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{F233A099-1241-ED48-AE78-14CEAAA87C95}\" hiddenColumns=\"1\"><selection pane=\"topLeft\" activeCell=\"F1\" sqref=\"F1:K1048576\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/></customSheetView><customSheetView guid=\"{16591214-8D55-8A47-82DC-2F7E1C7FCAA3}\" filter=\"1\" showAutoFilter=\"1\" hiddenColumns=\"1\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/><autoFilter ref=\"A1:K33\"><filterColumn colId=\"0\"><customFilters and=\"1\"><customFilter operator=\"greaterThan\" val=\"15\"/></customFilters></filterColumn></autoFilter></customSheetView><customSheetView guid=\"{AB8C9049-5EEE-4D41-AFBB-D05935ABA15C}\" filter=\"1\" showAutoFilter=\"1\" hiddenColumns=\"1\"><selection pane=\"topLeft\" activeCell=\"A1\" sqref=\"A1\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><printOptions gridLines=\"1\"/><pageSetup copies=\"1\" firstPageNumber=\"1\" fitToHeight=\"1\" fitToWidth=\"1\" horizontalDpi=\"300\" paperSize=\"9\" scale=\"100\" verticalDpi=\"300\"/><autoFilter ref=\"A1:K33\"><filterColumn colId=\"1\"><filters blank=\"0\"><filter val=\"4\"/></filters></filterColumn></autoFilter></customSheetView></customSheetViews>"
  got <- wb$worksheets[[1]]$customSheetViews
  expect_equal(exp, got)
})

test_that("xlsb formula line breaks are handled", {

  skip_online_checks()

  fl <- testfile_path("line_break.xlsb")

  fml <- "IF(A1 = 1,\"Value \n\"&A1,)"
  wb <- wb_workbook()$add_worksheet()$
    add_data(x = 1)$
    add_formula(x = fml, dims = "B1")$
    add_cell_style(dims = "B1", wrap_text = TRUE)$
    set_row_heights(rows = 1, heights = 30)

  wb2 <- wb_load(fl)
  exp <- "IF( A1= 1,\"Value \n\"&A1, )"
  fml2 <- wb2$to_df(show_formula = TRUE, col_names = FALSE)[1, "B"]
  expect_equal(exp,  fml2)

  fml  <- gsub(" ", "", fml)
  fml2 <- gsub(" ", "", fml2)

  expect_equal(fml, fml2)
})
