test_that("reading xls values works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_values.xls")
  wb <- wb_load(fl)

  expect_equal(unname(wb$get_sheet_names()), c("Data", 'Sp&c<e>s "q"', "Hidden"))

  df <- wb_to_df(wb, sheet = 1, col_names = FALSE, dims = "A1:A10")
  expect_equal(
    df$A,
    c(1, -5, 0.1, 123456.789, 1e20, 2147483647, -2147483648, 1 / 3, 0, 1234567.5)
  )

  df <- wb_to_df(wb, sheet = 1, col_names = FALSE)
  strs <- df$E[grepl("^s[0-9]{3}_", df$E)]
  expect_length(strs, 400)
  expect_true(all(nchar(strs) == 35))
  expect_equal(sum(nchar(unlist(df[sapply(df, is.character)])) == 9000, na.rm = TRUE), 1)
})

test_that("reading xls styles works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_styles.xls")
  wb <- wb_load(fl)

  expect_equal(wb$worksheets[[1]]$mergeCells, "<mergeCell ref=\"G1:H2\"/>")

  fills <- paste(wb$styles_mgr$styles$fills, collapse = "")
  expect_match(fills, "FFFFFF00")
  expect_match(fills, "FF99CCFF")
  expect_match(fills, "FFFF99CC")

  nf <- paste(wb$styles_mgr$styles$numFmts, collapse = "")
  expect_match(nf, "0.0%", fixed = TRUE)
  expect_match(nf, "EUR", fixed = TRUE)

  ra <- wb$worksheets[[1]]$sheet_data$row_attr
  expect_equal(as.numeric(ra[ra$r == "3", "ht"]), 40, tolerance = 0.01)
  expect_equal(ra[ra$r == "5", "hidden"], "1")
})

test_that("reading xls formulas works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_formulas.xls")
  wb <- wb_load(fl)

  f <- wb_to_df(wb, sheet = 1, col_names = FALSE, show_formula = TRUE, dims = "B1:F1")
  expect_equal(as.character(f$B), "A1*2+1")
  expect_equal(as.character(f$C), "SUM(A1:A5)")
  expect_equal(as.character(f$D), "IF(A1>2,\"big\",\"small\")")
  expect_equal(as.character(f$E), "Sheet2!A1+'My Sheet'!B2")
  expect_equal(as.character(f$F), "SUM(Sheet2:Sheet3!A1:A3)")

  v <- wb_to_df(wb, sheet = 1, col_names = FALSE, dims = "C1:F1")
  expect_equal(c(v$C, v$E, v$F), c(22.5, 103, 66))

  arr <- wb_to_df(wb, sheet = 1, col_names = FALSE, dims = "J1:J3")
  expect_equal(arr$J, c(3, 6, 9))

  dn <- paste(wb$workbook$definedNames, collapse = "")
  expect_match(dn, "Total", fixed = TRUE)
  expect_match(dn, "Rng", fixed = TRUE)
  expect_match(dn, "LocalTotal", fixed = TRUE)
  expect_match(dn, "localSheetId=\"0\"", fixed = TRUE)
})

test_that("reading an xls saved by Excel works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  expect_true(all(c("Values", "Styles", "Formulas", "Links", "Misc", "Hidden") %in% wb$get_sheet_names()))

  df <- wb_to_df(wb, sheet = 1)
  expect_equal(df$num, c(1, -5, 0.1, 123456.789, 1e20, 1 / 3))
  expect_equal(df$int, c(0L, 1L, -1L, 2147483647L, -2147483647L, 65536L))
  expect_equal(df$chr[2:5], c("b & <c>", "d \"q\"", "  lead", "ü€"))
  expect_equal(df$lgl, c(TRUE, FALSE, TRUE, FALSE, TRUE, FALSE))
  expect_equal(df$date, as.Date("2024-02-28") + 0:5)

  f <- wb_to_df(wb, sheet = 3, col_names = FALSE, show_formula = TRUE, dims = "B1:K1")
  expect_equal(
    unlist(f, use.names = FALSE),
    c(
      "A1*2+1", "SUM(A1:A5)", "IF(A1>2,\"big\",\"small\")", "Values!B2+Styles!A2",
      "SUM(Values!B2:B6)", "IFERROR(1/0,\"err\")", "VLOOKUP(2,A1:B5,2,FALSE)",
      "SUMPRODUCT(A1:A5,A1:A5)", "TEXT(TODAY(),\"yyyy\")", "COUNTIFS(A1:A5,\">1\",A1:A5,\"<5\")"
    )
  )

  filled <- wb_to_df(wb, sheet = 3, col_names = FALSE, show_formula = TRUE, dims = "O1:O20")
  expect_equal(as.character(filled$O), sprintf("A%d*2", 1:20))
})

test_that("reading xls hyperlinks, comments and validation works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  hl <- wb$worksheets[[4]]$hyperlinks
  expect_length(hl, 3)
  expect_match(hl[1], "ref=\"A1\"", fixed = TRUE)
  expect_match(hl[1], "tooltip=\"tip\"", fixed = TRUE)
  expect_match(hl[2], "location=\"Values!A1\"", fixed = TRUE)

  rels <- paste(wb$worksheets_rels[[4]], collapse = "")
  expect_match(rels, "Target=\"https://example.com/\"", fixed = TRUE)
  expect_match(rels, "Target=\"mailto:a@example.com\"", fixed = TRUE)

  cmt <- wb_get_comment(wb, sheet = 4)
  expect_true(cmt$ref %in% c("B1", "B2"))
  expect_match(paste(unlist(wb$comments), collapse = ""), "a comment", fixed = TRUE)

  dv <- wb$worksheets[[4]]$dataValidations
  expect_length(dv, 1)
  expect_match(dv, "type=\"whole\"", fixed = TRUE)
  expect_match(dv, "sqref=\"F1:F5\"", fixed = TRUE)
  expect_match(dv, "<formula1>1</formula1><formula2>10</formula2>", fixed = TRUE)
})

test_that("reading an xls table works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  tabs <- wb$tables
  skip_if(NROW(tabs) == 0L, "xls_source.xls has no table yet")

  expect_equal(tabs$tab_name, "tab")
  expect_equal(tabs$tab_ref, "D1:E4")
  expect_match(tabs$tab_xml, "<tableColumn id=\"1\" name=\"g\"/>", fixed = TRUE)
  expect_match(tabs$tab_xml, "<tableColumn id=\"2\" name=\"v\"/>", fixed = TRUE)
})

test_that("reading xls conditional formatting works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  cf <- wb$worksheets[[4]]$conditionalFormatting
  expect_equal(cf$sqref, "G1:G5")
  expect_match(cf$cf, "type=\"expression\"", fixed = TRUE)
  expect_match(cf$cf, "<formula>G1&gt;2</formula>", fixed = TRUE)

  # the formatting is a dxf with a font color and a fill
  dxfs <- wb$styles_mgr$styles$dxfs
  expect_length(dxfs, 1)
  expect_match(dxfs, "<font><color rgb=", fixed = TRUE)
  expect_match(dxfs, "<fill><patternFill", fixed = TRUE)
})

test_that("reading xls rich text works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  sst <- unlist(wb$sharedStrings)
  rich <- sst[grepl("<r>", sst, fixed = TRUE)]
  skip_if(length(rich) == 0L, "xls_source.xls has no rich text yet")

  expect_match(rich[1], "<b/>", fixed = TRUE)
  expect_match(rich[1], "FFFF0000", fixed = TRUE)
  expect_equal(wb_to_df(wb, "Misc", dims = "A1", col_names = FALSE)[[1]], "plain bold red")
})

test_that("reading xls page margins works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  margins <- wb$worksheets[[1]]$pageMargins
  expect_length(margins, 1)
  expect_match(margins, "<pageMargins left=\"[0-9.]+\" right=\"[0-9.]+\" top=\"[0-9.]+\" bottom=\"[0-9.]+\"")
})

test_that("reading xls color scales, data bars and icon sets works", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  wb <- wb_load(fl)

  cf <- wb$worksheets[[5]]$conditionalFormatting
  skip_if(length(cf) == 0L || !any(grepl("colorScale", cf$cf)), "xls_source.xls has no color scale yet")

  expect_true(any(grepl("<colorScale>", cf$cf, fixed = TRUE)))
  expect_true(any(grepl("<dataBar", cf$cf, fixed = TRUE)))
  expect_true(any(grepl("<iconSet iconSet=\"3Arrows\"", cf$cf, fixed = TRUE)))
})

test_that("pictures in xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  skip_if(length(wb$media) == 0L, "xls_source.xls has no pictures yet")
  expect_equal(length(wb$media), 2L)
  expect_true(any(grepl("xdr:pic", unlist(wb$drawings), fixed = TRUE)))
  tmp <- temp_xlsx()
  on.exit(unlink(tmp))
  wb$save(tmp)
  expect_equal(length(wb_load(tmp)$media), 2L)
})

test_that("loadExample.xls matches loadExample.xlsx", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  a <- wb_load(testfile_path("loadExample.xlsx"))
  b <- wb_load(testfile_path("loadExample.xls"))
  expect_equal(wb_get_sheet_names(b), wb_get_sheet_names(a))
  expect_equal(length(b$tables), length(a$tables))
  expect_equal(length(b$media), length(a$media))
  for (i in c(1, 3, 4)) {
    x <- wb_to_df(a, sheet = i, col_names = FALSE)
    y <- wb_to_df(b, sheet = i, col_names = FALSE)
    expect_equal(dim(y), dim(x))
    for (j in seq_along(x)) {
      xn <- suppressWarnings(as.numeric(x[[j]]))
      yn <- suppressWarnings(as.numeric(y[[j]]))
      num <- !is.na(xn) & !is.na(yn)
      expect_equal(yn[num], xn[num], tolerance = 1e-9)
      expect_equal(y[[j]][!num], x[[j]][!num])
    }
  }
})

test_that("charts in xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  charts <- unlist(wb$charts)
  expect_true(any(grepl("<c:scatterChart>", charts, fixed = TRUE)))
  expect_true(any(grepl("IrisSample!$D$2:$D$51", charts, fixed = TRUE)))
  expect_true(any(grepl("<c:barChart>", charts, fixed = TRUE)))
  expect_true(any(grepl("schemeClr val=\"accent1\"", charts, fixed = TRUE)))
  expect_true(any(grepl("<a:t>Petal</a:t>", charts, fixed = TRUE)))
  expect_true(any(grepl("<a:t> Width</a:t>", charts, fixed = TRUE)))

  drawings <- unlist(wb$drawings)
  expect_equal(sum(lengths(regmatches(drawings, gregexpr("<xdr:pic>", drawings, fixed = TRUE)))), 2L)
  expect_true(any(grepl("prst=\"smileyFace\"", drawings, fixed = TRUE)))

  tmp <- temp_xlsx()
  on.exit(unlink(tmp))
  wb$save(tmp)
  expect_equal(unlist(wb_load(tmp)$charts), charts)
})

test_that("chart sheets in xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  skip_if(!any(wb$is_chartsheet), "xls_source.xls has no chart sheet yet")
  expect_true(any(grepl("<c:chart", unlist(wb$charts), fixed = TRUE)))
})

test_that("document properties and workbook views of an xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  expect_match(wb$core, "<dc:creator>Alex</dc:creator>", fixed = TRUE)
  expect_match(wb$core, "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2014-03-07T16:08:25Z</dcterms:created>", fixed = TRUE)
  expect_match(wb$workbook$bookViews, "activeTab=\"1\"", fixed = TRUE)
})

test_that("sheet format and outline settings of an xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_match(wb$worksheets[[2]]$sheetFormatPr, "outlineLevelRow=\"1\"", fixed = TRUE)
})

test_that("named cell styles of an xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  cs <- wb$styles_mgr$styles$cellStyles
  expect_true(any(grepl("name=\"Normal\" xfId=\"0\" builtinId=\"0\"", cs, fixed = TRUE)))
  expect_true(any(grepl("name=\"Bad\".*builtinId=\"27\"", cs)))
  expect_gt(length(wb$styles_mgr$styles$cellStyleXfs), 1L)
})

test_that("sheet dimensions of an xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_equal(wb$worksheets[[1]]$dimension, "<dimension ref=\"A1:H7\"/>")
})

test_that("ignored errors of an xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  ie <- unlist(lapply(wb$worksheets, function(x) x$ignoredErrors))
  expect_match(ie, "sqref=\"G1:G3\" numberStoredAsText=\"1\"", fixed = TRUE)
  expect_match(ie, "sqref=\"H1:H2\" evalError=\"1\" emptyCellReference=\"1\"", fixed = TRUE)
})

test_that("a single-cell sheet of an xls has the dimension A1", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_equal(wb$worksheets[[8]]$dimension, "<dimension ref=\"A1\"/>")
})

test_that("sort state of an xls auto filter is imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_equal(
    wb$worksheets[[5]]$autoFilter,
    "<autoFilter ref=\"J1:K11\"><sortState ref=\"J2:K11\"><sortCondition descending=\"1\" ref=\"J2:J11\"/></sortState></autoFilter>"
  )
})

test_that("charts of an xls get cached values and the default data label block", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  chart <- paste(unlist(wb$charts), collapse = "")
  expect_match(chart, "<c:f>IrisSample!$D$2:$D$51</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"50\"/><c:pt idx=\"0\"><c:v>1.9</c:v>", fixed = TRUE)
  expect_match(chart, "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/>", fixed = TRUE)
})

test_that("legends, data bars and filter names of an xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  expect_match(paste(unlist(wb$charts), collapse = ""), "<c:legendPos val=\"r\"/><c:overlay", fixed = TRUE)

  wb <- wb_load(testfile_path("xls_source.xls"))
  expect_match(paste(unlist(wb$charts), collapse = ""), "<c:legendPos val=\"r\"/><c:layout><c:manualLayout>", fixed = TRUE)
  expect_false(any(grepl("showValue=\"0\"", unlist(lapply(wb$worksheets, function(x) x$conditionalFormatting)), fixed = TRUE)))
  expect_match(wb$workbook$definedNames, "_xlnm._FilterDatabase\" localSheetId=\"4\" hidden=\"1\">'Misc'!$J$1:$K$11", fixed = TRUE, all = FALSE)
})

test_that("data bars of an xls get the extension Excel uses for them", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("xls_source.xls"))
  ws <- wb$worksheets[[5]]
  expect_match(ws$conditionalFormatting$cf[grepl("dataBar", ws$conditionalFormatting$cf)], "<dataBar><cfvo type=\"min\"/>", fixed = TRUE)
  expect_match(ws$extLst, "<x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\"", fixed = TRUE)
})

test_that("data_only skips everything but the cells", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_source.xls")
  skip_if(!file.exists(fl))

  wb <- wb_load(fl, sheet = 3, data_only = TRUE)
  full <- wb_load(fl, sheet = 3)
  expect_equal(length(wb$charts), 0L)
  expect_equal(length(wb$tables), 0L)
  expect_message(a <- wb_to_df(wb, col_names = FALSE), "contains no data")
  expect_message(b <- wb_to_df(full, col_names = FALSE), "contains no data")
  expect_equal(a, b)
})

test_that("pivot tables are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_pivot.xls")
  skip_if(!file.exists(fl))

  wb <- wb_load(fl)
  expect_equal(length(wb$pivotTables), 2L)
  expect_equal(length(wb$pivotDefinitions), 1L)

  pt1 <- wb$pivotTables[[1]]
  expect_match(pt1, '<location ref="A3:C14"', fixed = TRUE)
  expect_match(pt1, '<rowFields count="2"><field x="0"/><field x="1"/></rowFields>', fixed = TRUE)
  expect_match(pt1, '<colFields count="1"><field x="-2"/></colFields>', fixed = TRUE)
  expect_match(pt1, '<pageField fld="4"', fixed = TRUE)
  expect_match(pt1, 'subtotal="average"', fixed = TRUE)

  def <- wb$pivotDefinitions[[1]]
  expect_match(def, '<worksheetSource ref="A1:E13" sheet="data"/>', fixed = TRUE)
  expect_match(def, '<cacheField name="Species"', fixed = TRUE)
  expect_match(def, '<s v="a"/><s v="b"/><s v="c"/>', fixed = TRUE)

  tmp <- temp_xlsx()
  wb$save(tmp)
  wb2 <- wb_load(tmp)
  expect_equal(length(wb2$pivotTables), 2L)

  expect_equal(length(wb_load(fl, data_only = TRUE, sheet = 1)$pivotTables), 0L)
})

test_that("shape numbers of comments do not overlap", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  fl <- testfile_path("xls_comments.xls")
  skip_if(!file.exists(fl))

  wb <- wb_load(fl)
  ids <- lapply(wb$vml, function(x) as.integer(regmatches(x, gregexpr("(?<=_x0000_s)[0-9]+", x, perl = TRUE))[[1]]))
  expect_equal(sum(duplicated(unlist(ids))), 0L)
  expect_equal(lengths(ids), lengths(wb$comments))
  expect_match(wb$vml[[1]], 'data="1,2"', fixed = TRUE)
})

test_that("xls paths with a tilde are expanded", {
  skip_on_cran()
  home <- tempfile()
  dir.create(home)
  file.copy(testfile_path("loadExample.xls"), file.path(home, "loadExample.xls"))
  old <- Sys.getenv("HOME")
  Sys.setenv(HOME = home)
  on.exit(Sys.setenv(HOME = old))
  wb <- wb_load("~/loadExample.xls")
  expect_equal(length(wb$get_sheet_names()), 4L)
})

test_that("conditional formatting of xls uses the full colors", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  dxfs <- wb$styles_mgr$styles$dxfs
  expect_match(dxfs, "FFFFC7CE", fixed = TRUE, all = FALSE)
  expect_match(dxfs, "FF9C0006", fixed = TRUE, all = FALSE)
})

test_that("pivot table style of xls is imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  a <- wb_load(testfile_path("loadExample.xlsx"))
  b <- wb_load(testfile_path("loadExample.xls"))
  style <- function(wb) regmatches(wb$pivotTables, regexpr("<pivotTableStyleInfo[^>]*>", wb$pivotTables))
  expect_equal(style(b), style(a))
})

test_that("pivot field options of xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  a <- wb_load(testfile_path("xls_cov2_pivots.xlsx"))
  b <- wb_load(testfile_path("xls_cov2_pivots.xls"))
  attrs <- function(wb, i, what) {
    x <- unlist(regmatches(wb$pivotTables[i], gregexpr("<pivotField\\b[^>]*>", wb$pivotTables[i])))
    sort(unlist(regmatches(x, gregexpr(paste0(what, '="[^"]*"'), x))))
  }
  for (i in 10:11) expect_equal(attrs(b, i, "sortType"), attrs(a, i, "sortType"))
  expect_equal(attrs(b, 13, "subtotalTop"), attrs(a, 13, "subtotalTop"))
  for (w in c("dragToRow", "dragToCol", "dragToPage")) expect_equal(attrs(b, 20, w), attrs(a, 20, w))
  expect_length(attrs(b, 10, "sortType"), 3L)
})

test_that("shared string counts of xls are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  wb <- wb_load(testfile_path("loadExample.xls"))
  expect_equal(attr(wb$sharedStrings, "uniqueCount"), as.character(length(wb$sharedStrings)))
})

test_that("page layout view of xls is imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  fl <- testfile_path("xls_source.xls")
  raw_xls <- readBin(fl, "raw", file.size(fl))
  plv <- as.raw(c(0x8b, 0x08, 0x10, 0x00, 0x8b, 0x08))
  at <- which(raw_xls == plv[1])
  at <- at[vapply(at, function(i) all(raw_xls[i + 0:5] == plv), NA)][1]
  raw_xls[at + 16] <- as.raw(80)
  raw_xls[at + 18] <- as.raw(3)
  tmp <- tempfile(fileext = ".xls")
  on.exit(unlink(tmp), add = TRUE)
  writeBin(raw_xls, tmp)
  wb <- wb_load(tmp)
  views <- unlist(lapply(wb$worksheets, function(x) x$sheetViews))
  expect_equal(sum(grepl('view="pageLayout"', views, fixed = TRUE)), 1L)
  expect_true(any(grepl('zoomScalePageLayoutView="80"', views, fixed = TRUE)))
})

test_that("page break preview of xls is imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()
  fl <- testfile_path("xls_source.xls")
  raw_xls <- readBin(fl, "raw", file.size(fl))
  w2 <- as.raw(c(0x3e, 0x02, 0x12, 0x00))
  at <- which(raw_xls == w2[1])
  at <- at[vapply(at, function(i) all(raw_xls[i + 0:3] == w2), NA)][1]
  raw_xls[at + 5] <- as.raw(bitwOr(as.integer(raw_xls[at + 5]), 0x08L))
  raw_xls[at + 14] <- as.raw(60)
  raw_xls[at + 15] <- as.raw(0)
  tmp <- tempfile(fileext = ".xls")
  on.exit(unlink(tmp), add = TRUE)
  writeBin(raw_xls, tmp)
  wb <- wb_load(tmp)
  views <- unlist(lapply(wb$worksheets, function(x) x$sheetViews))
  expect_true(any(grepl('view="pageBreakPreview"', views, fixed = TRUE)))
  expect_true(any(grepl('zoomScaleSheetLayoutView="60"', views, fixed = TRUE)))
  expect_true(any(grepl('zoomScale="60"', views, fixed = TRUE)))
})

test_that("rels of chart sheets are read if no worksheet has rels", {
  skip_on_cran()
  tmp <- tempfile("chartsheet_rels")
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  wb <- wb_workbook()$add_worksheet()$add_data(x = 1)$add_chartsheet()
  fl <- tempfile(fileext = ".xlsx")
  on.exit(unlink(fl), add = TRUE)
  wb$save(fl)
  fls <- utils::unzip(fl, exdir = tmp)
  dir.create(file.path(tmp, "xl", "chartsheets", "_rels"), recursive = TRUE, showWarnings = FALSE)
  dir.create(file.path(tmp, "xl", "drawings"), showWarnings = FALSE)
  sheet <- grep("chartsheets/sheet[0-9]+[.]xml$", fls, value = TRUE)
  xml <- paste(readLines(sheet, warn = FALSE), collapse = "")
  xml <- sub("</chartsheet>", "<drawing r:id=\"rId1\"/></chartsheet>", xml, fixed = TRUE)
  if (!grepl("xmlns:r=", xml, fixed = TRUE))
    xml <- sub("<chartsheet ", "<chartsheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ", xml, fixed = TRUE)
  writeLines(xml, sheet)
  writeLines(paste0(
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>",
    "</Relationships>"
  ), file.path(tmp, "xl", "chartsheets", "_rels", basename(paste0(sheet, ".rels"))))
  writeLines(
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>",
    file.path(tmp, "xl", "drawings", "drawing1.xml")
  )
  out <- tempfile(fileext = ".xlsx")
  on.exit(unlink(out), add = TRUE)
  old <- setwd(tmp)
  zip::zip(out, files = list.files(tmp, recursive = TRUE, all.files = TRUE), mode = "mirror")
  setwd(old)

  wb2 <- wb_load(out)
  expect_true(any(grepl("drawing1.xml", unlist(wb2$worksheets_rels), fixed = TRUE)))
})

test_that("xls files created from the xls_cov workbooks are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()

  expected <- list(
    xls_cov_views = c("freeze", "freeze_row", "freeze_col", "split", "hidden_outline", "protected", "very_hidden"),
    xls_cov_features = c("filter", "merge_val", "cf", "links", "page", "fit"),
    xls_cov_names_formulas = c("data", "calc"),
    xls_cov_drawings = c("img", "chart", "chartsheet"),
    xls_cov_pivot_external = c("src", "pivot", "pivot2", "ext")
  )

  for (nm in names(expected)) {
    fl <- testfile_path(paste0(nm, ".xls"))

    wb <- wb_load(fl)
    expect_equal(wb$get_sheet_names(escape = TRUE), expected[[nm]], ignore_attr = TRUE, info = nm)

    tmp <- tempfile(fileext = ".xlsx")
    wb$save(tmp)
    wb2 <- wb_load(tmp)
    expect_equal(length(wb2$worksheets), length(expected[[nm]]), info = nm)
    unlink(tmp)
  }

  wb <- wb_load(testfile_path("xls_cov_views.xls"))
  panes <- unlist(lapply(wb$worksheets, function(x) x$freezePane))
  expect_equal(sum(grepl('state="frozen"', panes, fixed = TRUE)), 3L)
  expect_equal(sum(grepl('state="split"', panes, fixed = TRUE)), 1L)

  wb <- wb_load(testfile_path("xls_cov_features.xls"))
  expect_true(length(wb$worksheets[[1]]$autoFilter) > 0)
  expect_true(length(wb$worksheets[[2]]$mergeCells) > 0)
  expect_true(length(wb$worksheets[[2]]$dataValidations) > 0)
  expect_true(length(wb$worksheets[[3]]$conditionalFormatting) > 0)
  cf_txt <- unlist(wb$worksheets[[3]]$conditionalFormatting)
  expect_match(cf_txt, 'type="containsText" operator="containsText" text="Mazda"', fixed = TRUE, all = FALSE)
  expect_match(cf_txt, 'type="endsWith" operator="endsWith" text="4"', fixed = TRUE, all = FALSE)
  expect_true(length(wb$worksheets[[4]]$hyperlinks) > 0)

  wb <- wb_load(testfile_path("xls_cov_names_formulas.xls"))
  expect_true(nrow(wb_get_named_regions(wb)) > 0)
  fml <- wb_to_df(wb, sheet = "calc", col_names = FALSE, show_formula = TRUE)
  expect_true(sum(grepl("^[A-Z]", unlist(fml[[3]]))) > 50)

  wb <- wb_load(testfile_path("xls_cov_drawings.xls"))
  expect_true(length(wb$charts) > 0)
  expect_true(length(wb$media) > 0)
})

test_that("xls files created from the xls_cov2 workbooks are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()

  wb <- wb_load(testfile_path("xls_cov2_pivots.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE)[1:3], c("src", "functions", "functions2"), ignore_attr = TRUE)
  expect_true(length(wb$pivotTables) >= 20)
  expect_true(any(grepl("defaultSubtotal=\"0\"", wb$pivotTables, fixed = TRUE)))
  # a table with column fields only has no row fields
  expect_true(any(grepl("<colFields", wb$pivotTables, fixed = TRUE) & !grepl("<rowFields", wb$pivotTables, fixed = TRUE)))
  # the lines of the table match the area of the table
  expect_true(any(grepl("<rowItems count=\"11\">", wb$pivotTables, fixed = TRUE)))
  # data fields on the axis are numbered, dates of the cache are kept
  expect_true(any(grepl("<i i=\"3\">", wb$pivotTables, fixed = TRUE)))
  expect_true(any(grepl("<d v=\"2024-01-01T00:00:00\"/>", wb$pivotDefinitions, fixed = TRUE)))
  tmp <- tempfile(fileext = ".xlsx")
  wb$save(tmp)
  expect_equal(length(wb_load(tmp)$pivotTables), length(wb$pivotTables))
  unlink(tmp)

  wb <- wb_load(testfile_path("xls_cov2_values.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE), c("errors", "numfmts", "numbers", "colors", "fonts"), ignore_attr = TRUE)
  cc <- wb$worksheets[[1]]$sheet_data$cc
  errs <- cc$v[cc$c_t == "e"]
  expect_true(all(c("#N/A", "#DIV/0!", "#VALUE!", "#NUM!", "#NAME?", "#NULL!", "#REF!") %in% errs))
  expect_equal(nrow(wb_to_df(wb, sheet = "numfmts", col_names = FALSE)), 32L)

  wb <- wb_load(testfile_path("xls_cov2_names.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE), c("a", "b", "c with space"), ignore_attr = TRUE)
  expect_true(nrow(wb_get_named_regions(wb)) > 10)

  wb <- wb_load(testfile_path("xls_cov2_shapes.xls"))
  expect_equal(wb$get_sheet_names(escape = TRUE), c("shapes", "textboxes", "connectors"), ignore_attr = TRUE)
  expect_true(sum(lengths(regmatches(wb$drawings, gregexpr("<xdr:sp[ >]", wb$drawings)))) > 70)
  expect_equal(sum(lengths(regmatches(wb$drawings, gregexpr("<xdr:cxnSp", wb$drawings)))), 6L)
  expect_equal(sum(lengths(regmatches(wb$drawings, gregexpr("<xdr:grpSp>", wb$drawings)))), 6L)
  expect_true(any(grepl("plain text box", wb$drawings, fixed = TRUE)))
  expect_true(any(grepl("<a:prstDash val=\"sysDot\"/>", wb$drawings, fixed = TRUE)))
})

test_that("a single sheet of an xls file can be read without converting the others", {
  fl <- testfile_path("xls_cov2_values.xls")
  skip_if(fl == "" || !file.exists(fl))

  full <- wb_load(fl)
  nms <- full$get_sheet_names()

  wb <- wb_load(fl, sheet = 3, data_only = TRUE)
  expect_equal(wb_to_df(wb, sheet = 3), wb_to_df(full, sheet = 3))
  expect_null(suppressMessages(suppressWarnings(wb_to_df(wb, sheet = 1))))

  wb <- wb_load(fl, sheet = names(nms)[2], data_only = TRUE)
  expect_equal(wb_to_df(wb, sheet = 2), wb_to_df(full, sheet = 2))

  expect_error(wb_load(fl, sheet = "nope", data_only = TRUE), "No such sheet")
})

test_that("single sheets of xls files keep their strings, dates and names", {
  for (file in c("xls_values.xls", "xls_cov2_values.xls", "loadExample.xls")) {
    fl <- testfile_path(file)
    skip_if(fl == "" || !file.exists(fl))

    full <- wb_load(fl)
    nms <- full$get_sheet_names()

    for (i in seq_along(nms)) {
      if (full$is_chartsheet[i]) next
      expected <- wb_to_df(full, sheet = i)

      by_index <- wb_load(fl, sheet = i, data_only = TRUE)
      expect_equal(wb_to_df(by_index, sheet = i), expected)

      by_name <- wb_load(fl, sheet = names(nms)[i], data_only = TRUE)
      expect_equal(wb_to_df(by_name, sheet = i), expected)

      expect_equal(by_index$get_sheet_names(), nms)
    }
  }

  fl <- testfile_path("xls_cov2_values.xls")
  skip_if(fl == "" || !file.exists(fl))
  df <- wb_to_df(wb_load(fl, sheet = 3, data_only = TRUE), sheet = 3)
  expect_true(any(vapply(df, is.character, NA)))
  expect_true(any(vapply(df, inherits, NA, c("Date", "POSIXct"))))
})

test_that("xls_cov_values with comments is imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()

  fl <- testfile_path("xls_cov_values.xls")

  wb <- wb_load(fl)
  expect_equal(wb$get_sheet_names(escape = TRUE), "values", ignore_attr = TRUE)
  df <- wb_to_df(wb, sheet = "values")
  expect_true(nrow(df) >= 9)
  expect_equal(nchar(df[[1]][3]), 32767L)

  wb2 <- wb_load(fl, sheet = 1, data_only = TRUE)
  expect_equal(wb_to_df(wb2, sheet = 1), df)
})

test_that("xls files created from the xls_cov3 and xls_cov4 workbooks are imported", {
  skip_if(.Platform$endian == "big", "XLS: skipping on big-endian system")
  skip_on_cran()

  for (nm in c("xls_cov3_cf", "xls_cov4_misc")) {
    wb_s <- suppressWarnings(wb_load(testfile_path(paste0(nm, ".xls"))))
    wb_x <- wb_load(testfile_path(paste0(nm, ".xlsx")))
    expect_equal(wb_s$get_sheet_names(), wb_x$get_sheet_names(), info = nm)
    expect_equal(length(wb_s$tables$tab_xml), length(wb_x$tables$tab_xml), info = nm)

    tmp <- tempfile(fileext = ".xlsx")
    suppressWarnings(wb_s$save(tmp))
    expect_equal(length(wb_load(tmp)$worksheets), length(wb_x$worksheets), info = nm)
    unlink(tmp)
  }

  wb_s <- wb_load(testfile_path("xls_cov4_misc.xls"))
  expect_length(wb_s$externalLinks, 1L)

  cf <- function(f, i) paste(unlist(suppressWarnings(wb_load(testfile_path(f)))$worksheets[[i]]$conditionalFormatting[2]), collapse = "")
  cf_s <- vapply(1:5, function(i) cf("xls_cov3_cf.xls", i), "")
  expect_match(cf_s[1], "type=\"timePeriod\" timePeriod=\"lastWeek\"", fixed = TRUE)
  expect_match(cf_s[2], "type=\"aboveAverage\" equalAverage=\"1\"", fixed = TRUE)
  expect_match(cf_s[2], "stdDev=\"3\"", fixed = TRUE)
  expect_match(cf_s[3], "iconSet=\"3TrafficLights2\" reverse=\"1\"", fixed = TRUE)
  expect_match(cf_s[5], "type=\"top10\" rank=\"3\" percent=\"0\" bottom=\"1\"", fixed = TRUE)
  expect_match(cf_s[5], "stopIfTrue=\"1\"", fixed = TRUE)

  tab <- wb_s$tables$tab_xml
  expect_match(tab, "<filter val=\"alpha\"/>", fixed = TRUE)
  expect_match(tab, "<customFilter operator=\"greaterThan\" val=\"10\"/>", fixed = TRUE)
  expect_match(tab, "<dynamicFilter type=\"aboveAverage\" val=\"50\"/>", fixed = TRUE)
  expect_match(tab, "<colorFilter dxfId=\"0\" cellColor=\"1\"/>", fixed = TRUE)
  expect_match(tab, "<iconFilter iconSet=\"3TrafficLights1\" iconId=\"0\"/>", fixed = TRUE)
  expect_match(wb_s$styles_mgr$styles$dxfs[1], "FFFF0000", fixed = TRUE)

  af <- wb_s$worksheets[[2]]$autoFilter
  expect_match(af, "<filters blank=\"1\"><filter val=\"3\"/><filter val=\"7\"/></filters>", fixed = TRUE)
  expect_match(af, "<dateGroupItem year=\"2020\" month=\"3\" day=\"5\" dateTimeGrouping=\"day\"/>", fixed = TRUE)
  expect_match(af, "<dynamicFilter type=\"belowAverage\" val=\"10.5\"/>", fixed = TRUE)
})
