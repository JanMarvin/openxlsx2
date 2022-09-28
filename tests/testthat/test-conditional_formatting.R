my_workbook <- function() {
  # Make the workbook each time so wbWorkbook$debug() is easier to access within
  # each test
  wb <- wb_workbook()
  negStyle <- create_dxfs_style(font_color = wb_colour(hex = "FF9C0006"), bgFill = wb_colour(hex = "FFFFC7CE"))
  wb$styles_mgr$add(negStyle, "negStyle")
  posStyle <- create_dxfs_style(font_color = wb_colour(hex = "FF006100"), bgFill = wb_colour(hex = "FFC6EFCE"))
  wb$styles_mgr$add(posStyle, "posStyle")
  invisible(wb)
}

expect_save <- function(wb) {
  # Quick function to check original saving.
  assert_workbook(wb)
  path <- temp_xlsx()

  testthat::expect_silent(wb_save(wb, path))
  testthat::expect_silent(wb1 <- wb_load(path))

  for (sheet in seq_along(wb$sheet_names))
    testthat::expect_identical(
      read_xml(wb$worksheets[[sheet]]$conditionalFormatting, pointer = FALSE),
      read_xml(wb1$worksheets[[sheet]]$conditionalFormatting, pointer = FALSE)
    )

  invisible()
}

test_that("type = 'expression' work", {
  wb <- my_workbook()
  wb$add_worksheet("cellIs")

  exp <- c(
    '<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
    '<dxf><font><color rgb="FF006100"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>'
  )

  expect_identical(exp, wb$styles_mgr$styles$dxfs)

  set.seed(123)

  ## rule applies to all each cell in range
  wb$add_data("cellIs", -5:5)
  wb$add_data("cellIs", LETTERS[1:11], startCol = 2)
  wb$add_conditional_formatting(
    sheet = "cellIs",
    cols = 1,
    rows = 1:11,
    rule = "!=0",
    style = "negStyle"
  )
  wb$add_conditional_formatting(
    sheet = "cellIs",
    cols = 1,
    rows = 1:11,
    rule = "==0",
    style = "posStyle"
  )

  exp <- c(
    `A1:A11` = '<cfRule type="expression" dxfId="0" priority="2"><formula>A1&lt;&gt;0</formula></cfRule>',
    `A1:A11` = '<cfRule type="expression" dxfId="1" priority="1"><formula>A1=0</formula></cfRule>'
  )

  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)


  wb$add_worksheet("Moving Row")
  ## highlight row dependent on first cell in row
  wb$add_data("Moving Row", -5:5)
  wb$add_data("Moving Row", LETTERS[1:11], startCol = 2)
  wb$add_conditional_formatting(
    "Moving Row",
    cols = 1:2,
    rows = 1:11,
    rule = "$A1<0",
    style = "negStyle"
  )
  wb$add_conditional_formatting(
    "Moving Row",
    cols = 1:2,
    rows = 1:11,
    rule = "$A1>0",
    style = "posStyle"
  )

  exp <- c(
    `A1:B11` = '<cfRule type="expression" dxfId="0" priority="2"><formula>$A1&lt;0</formula></cfRule>',
    `A1:B11` = '<cfRule type="expression" dxfId="1" priority="1"><formula>$A1&gt;0</formula></cfRule>'
  )
  expect_identical(exp, wb$worksheets[[2]]$conditionalFormatting)

  wb$add_worksheet("Moving Col")
  ## highlight column dependent on first cell in column
  wb$add_data("Moving Col", -5:5)
  wb$add_data("Moving Col", LETTERS[1:11], startCol = 2)
  wb$add_conditional_formatting(
    "Moving Col",
    cols = 1:2,
    rows = 1:11,
    rule = "A$1<0",
    style = "negStyle"
  )
  wb$add_conditional_formatting(
    "Moving Col",
    cols = 1:2,
    rows = 1:11,
    rule = "A$1>0",
    style = "posStyle"
  )

  exp <- c(
    `A1:B11` = '<cfRule type="expression" dxfId="0" priority="2"><formula>A$1&lt;0</formula></cfRule>',
    `A1:B11` = '<cfRule type="expression" dxfId="1" priority="1"><formula>A$1&gt;0</formula></cfRule>'
  )
  expect_identical(exp, wb$worksheets[[3]]$conditionalFormatting)

  wb$add_worksheet("Dependent on")
  ## highlight entire range cols X rows dependent only on cell A1
  wb$add_data("Dependent on", -5:5)
  wb$add_data("Dependent on", LETTERS[1:11], startCol = 2)
  wb$add_conditional_formatting(
    "Dependent on",
    cols = 1:2,
    rows = 1:11,
    rule = "$A$1<0",
    style = "negStyle"
  )
  wb$add_conditional_formatting(
    "Dependent on",
    cols = 1:2,
    rows = 1:11,
    rule = "$A$1>0",
    style = "posStyle"
  )

  exp <- c(
    `A1:B11` = '<cfRule type="expression" dxfId="0" priority="2"><formula>$A$1&lt;0</formula></cfRule>',
    `A1:B11` = '<cfRule type="expression" dxfId="1" priority="1"><formula>$A$1&gt;0</formula></cfRule>'
  )
  expect_identical(exp, wb$worksheets[[4]]$conditionalFormatting)


  ## highlight cells in column 1 based on value in column 2
  wb$add_data("Dependent on", data.frame(x = 1:10, y = runif(10)), startRow = 15)
  wb$add_conditional_formatting(
    "Dependent on",
    cols = 1,
    rows = 16:25,
    rule = "B16<0.5",
    style = "negStyle"
  )
  wb$add_conditional_formatting(
    "Dependent on",
    cols = 1,
    rows = 16:25,
    rule = "B16>=0.5",
    style = "posStyle"
  )

  exp <- c(
    `A1:B11`  = '<cfRule type="expression" dxfId="0" priority="4"><formula>$A$1&lt;0</formula></cfRule>',
    `A1:B11`  = '<cfRule type="expression" dxfId="1" priority="3"><formula>$A$1&gt;0</formula></cfRule>',
    `A16:A25` = '<cfRule type="expression" dxfId="0" priority="2"><formula>B16&lt;0.5</formula></cfRule>',
    `A16:A25` = '<cfRule type="expression" dxfId="1" priority="1"><formula>B16&gt;=0.5</formula></cfRule>'
  )
  expect_identical(exp, wb$worksheets[[4]]$conditionalFormatting)

  expect_save(wb)
})

test_that("type = 'duplicated' works", {
  wb <- my_workbook()
  wb$add_worksheet("Duplicates")
  ## highlight duplicates using default style
  wb$add_data("Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
  wb$add_conditional_formatting("Duplicates", cols = 1, rows = 1:10, type = "duplicated")

  exp <- c(`A1:A10` = '<cfRule type="duplicateValues" dxfId="2" priority="1"/>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'containsText' works", {
  wb <- my_workbook()
  wb$add_worksheet("containsText")
  ## cells containing text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("containsText", sapply(1:10, fn))
  wb$add_conditional_formatting(
    "containsText",
    cols = 1,
    rows = 1:10,
    type = "containsText",
    rule = "A"
  )

  # TODO remove identing from xml
  exp <- c(`A1:A10` = '<cfRule type="containsText" dxfId="2" priority="1" operator="containsText" text="A"><formula>NOT(ISERROR(SEARCH("A", A1)))</formula></cfRule>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'notContainsText' works", {
  wb <- my_workbook()
  wb$add_worksheet("notcontainsText")
  ## cells not containing text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("notcontainsText", x = sapply(1:10, fn))
  wb$add_conditional_formatting(
    "notcontainsText",
    cols = 1,
    rows = 1:10,
    type = "notContainsText",
    rule = "A"
  )

  exp <- c(`A1:A10` = '<cfRule type="notContainsText" dxfId="2" priority="1" operator="notContains" text="A"><formula>ISERROR(SEARCH("A", A1))</formula></cfRule>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'beginsWith' works", {
  wb <- my_workbook()
  wb$add_worksheet("beginsWith")
  ## cells begins with text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("beginsWith", sapply(1:100, fn))
  wb$add_conditional_formatting("beginsWith", cols = 1, rows = 1:100, type = "beginsWith", rule = "A")

  exp <- c(`A1:A100` = '<cfRule type="beginsWith" dxfId="2" priority="1" operator="beginsWith" text="A"><formula>LEFT(A1,LEN("A"))="A"</formula></cfRule>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'endsWith' works", {
  wb <- my_workbook()
  wb$add_worksheet("endsWith")
  ## cells ends with text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("endsWith", sapply(1:100, fn))
  wb$add_conditional_formatting(
    "endsWith",
    cols = 1,
    rows = 1:100,
    type = "endsWith",
    rule = "A"
  )

  exp <- c(`A1:A100` = '<cfRule type="endsWith" dxfId="2" priority="1" operator="endsWith" text="A"><formula>RIGHT(A1,LEN("A"))="A"</formula></cfRule>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'colorScale' works", {
  wb <- my_workbook()
  wb$add_worksheet("colourScale", zoom = 30)
  ## colourscale colours cells based on cell value
  df <- read_xlsx(system.file("extdata", "readTest.xlsx", package = "openxlsx2"), sheet = 4)
  wb$add_data("colourScale", df, colNames = FALSE) ## write data.frame
  ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
  ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
  wb$add_conditional_formatting(
    "colourScale",
    cols = seq_along(df),
    rows = seq_len(nrow(df)),
    style = c("black", "white"),
    rule = c(0, 255),
    type = "colorScale"
  )
  wb$set_col_widths("colourScale", cols = seq_along(df), widths = 1.07)
  wb <- wb_set_row_heights(wb, "colourScale", rows = seq_len(nrow(df)), heights = 7.5)

  exp <- c(`A1:E5` = '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="0"/><cfvo type="num" val="255"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'databar' works", {
  wb <- my_workbook()
  wb$add_worksheet("databar")
  ## Databars
  wb$add_data("databar", -5:5)
  wb$add_conditional_formatting("databar", cols = 1, rows = 1:11, type = "dataBar") ## Default colours

  exp <- c(`A1:A11` = '<cfRule type="dataBar" priority="1"><dataBar showValue="1"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40000}</x14:id></ext></extLst></cfRule>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'between' works", {
  wb <- my_workbook()
  wb$add_worksheet("between")
  ## Between
  # Highlight cells in interval [-2, 2]
  wb$add_data("between", -5:5)
  wb$add_conditional_formatting(
    "between",
    cols = 1,
    rows = 1:11,
    type = "between",
    rule = c(-2, 2)
  )

  exp <- c(`A1:A11` = '<cfRule type="cellIs" dxfId="2" priority="1" operator="between"><formula>-2</formula><formula>2</formula></cfRule>')
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'topN' works", {
  wb <- my_workbook()
  wb$add_worksheet("topN")
  ## Top N
  wb$add_data("topN", data.frame(x = 1:10, y = rnorm(10)))
  # Highlight top 5 values in column x
  wb$add_conditional_formatting(
    "topN",
    cols = 1,
    rows = 2:11,
    style = "posStyle",
    type = "topN",
    params = list(rank = 5)
  )

  # Highlight top 20 percentage in column y
  wb$add_conditional_formatting(
    "topN",
    cols = 2,
    rows = 2:11,
    style = "posStyle",
    type = "topN",
    params = list(rank = 20, percent = TRUE)
  )

  exp <- c(
    `A2:A11` = '<cfRule type="top10" dxfId="1" priority="2" rank="5" percent="0"/>',
    `B2:B11` = '<cfRule type="top10" dxfId="1" priority="1" rank="20" percent="1"/>'
  )
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type = 'bottomN' works", {
  wb <- my_workbook()
  wb$add_worksheet("bottomN")
  ## Bottom N
  wb$add_data("bottomN", data.frame(x = 1:10, y = rnorm(10)))
  # Highlight bottom 5 values in column x
  wb$add_conditional_formatting(
    "bottomN",
    cols = 1,
    rows = 2:11,
    style = "negStyle",
    type = "bottomN",
    params = list(rank = 5)
  )

  # Highlight bottom 20 percentage in column y
  wb$add_conditional_formatting(
    "bottomN",
    cols = 2,
    rows = 2:11,
    style = "negStyle",
    type = "bottomN",
    params = list(rank = 20, percent = TRUE)
  )

  exp <- c(
    `A2:A11` = "<cfRule type=\"top10\" dxfId=\"0\" priority=\"2\" rank=\"5\" percent=\"0\" bottom=\"1\"/>",
    `B2:B11` = "<cfRule type=\"top10\" dxfId=\"0\" priority=\"1\" rank=\"20\" percent=\"1\" bottom=\"1\"/>"
  )

  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("type as logical operators work", {
  wb <- my_workbook()
  wb$add_worksheet("logical operators")
  ## Logical Operators
  # You can use Excels logical Operators
  wb$add_data("logical operators", 1:10)
  wb$add_conditional_formatting(
    "logical operators",
    cols = 1,
    rows = 1:10,
    rule = "OR($A1=1,$A1=3,$A1=5,$A1=7)"
  )

  exp <- c(`A1:A10` = "<cfRule type=\"expression\" dxfId=\"2\" priority=\"1\"><formula>OR($A1=1,$A1=3,$A1=5,$A1=7)</formula></cfRule>")
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)
  expect_save(wb)
})

test_that("colorScale", {

  wb <- wb_workbook()

  ### two colors
  wb$add_worksheet("colourScale1", zoom = 30)
  ## colourscale colours cells based on cell value
  df <- read_xlsx(system.file("extdata", "readTest.xlsx", package = "openxlsx2"), sheet = 5)
  wb$add_data("colourScale1", df, colNames = FALSE) ## write data.frame
  ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
  ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
  wb$add_conditional_formatting(
    "colourScale1",
    cols = seq_along(df),
    rows = seq_len(nrow(df)),
    style = c("black", "white"),
    type = "colorScale"
  )
  wb$set_col_widths("colourScale1", cols = seq_along(df), widths = 2)
  wb <- wb_set_row_heights(wb, "colourScale1", rows = seq_len(nrow(df)), heights = 7.5)

  exp <- c(`A1:KK271` = "<cfRule type=\"colorScale\" priority=\"1\"><colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF000000\"/><color rgb=\"FFFFFFFF\"/></colorScale></cfRule>")
  expect_identical(exp, wb$worksheets[[1]]$conditionalFormatting)

  ### two colors and rule
  wb$add_worksheet("colourScale2", zoom = 30)
  wb$add_data("colourScale2", df, colNames = FALSE) ## write data.frame
  ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
  ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
  wb$add_conditional_formatting(
    "colourScale2",
    cols = seq_along(df),
    rows = seq_len(nrow(df)),
    style = c("blue", "red"),
    rule = c(1, 255),
    type = "colorScale"
  )
  wb$set_col_widths("colourScale2", cols = seq_along(df), widths = 2)
  wb <- wb_set_row_heights(wb, "colourScale2", rows = seq_len(nrow(df)), heights = 7.5)

  exp <- c(`A1:KK271` = "<cfRule type=\"colorScale\" priority=\"1\"><colorScale><cfvo type=\"num\" val=\"1\"/><cfvo type=\"num\" val=\"255\"/><color rgb=\"FF0000FF\"/><color rgb=\"FFFF0000\"/></colorScale></cfRule>")
  expect_identical(exp, wb$worksheets[[2]]$conditionalFormatting)

  ### three colors
  wb$add_worksheet("colourScale3", zoom = 30)
  wb$add_data("colourScale3", df, colNames = FALSE) ## write data.frame
  ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
  ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
  wb$add_conditional_formatting(
    "colourScale3",
    cols = seq_along(df),
    rows = seq_len(nrow(df)),
    style = c("red", "green", "blue"),
    type = "colorScale"
  )
  wb$set_col_widths("colourScale3", cols = seq_along(df), widths = 2)
  wb <- wb_set_row_heights(wb, "colourScale3", rows = seq_len(nrow(df)), heights = 7.5)

  exp <- c(`A1:KK271` = "<cfRule type=\"colorScale\" priority=\"1\"><colorScale><cfvo type=\"min\"/><cfvo type=\"percentile\" val=\"50\"/><cfvo type=\"max\"/><color rgb=\"FFFF0000\"/><color rgb=\"FF00FF00\"/><color rgb=\"FF0000FF\"/></colorScale></cfRule>")
  expect_identical(exp, wb$worksheets[[3]]$conditionalFormatting)

  wb$set_col_widths("colourScale3", cols = seq_along(df), widths = 2)
  wb <- wb_set_row_heights(wb, "colourScale3", rows = seq_len(nrow(df)), heights = 7.5)

  ### three colors and rule
  wb$add_worksheet("colourScale4", zoom = 30)
  wb$add_data("colourScale4", df, colNames = FALSE) ## write data.frame
  ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
  ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
  wb$add_conditional_formatting(
    "colourScale4",
    cols = seq_along(df),
    rows = seq_len(nrow(df)),
    style = c("red", "green", "blue"),
    rule = c(1, 155, 255),
    type = "colorScale"
  )
  wb$set_col_widths("colourScale4", cols = seq_along(df), widths = 2)
  wb <- wb_set_row_heights(wb, "colourScale4", rows = seq_len(nrow(df)), heights = 7.5)

  exp <- c(`A1:KK271` = "<cfRule type=\"colorScale\" priority=\"1\"><colorScale><cfvo type=\"num\" val=\"1\"/><cfvo type=\"num\" val=\"155\"/><cfvo type=\"num\" val=\"255\"/><color rgb=\"FFFF0000\"/><color rgb=\"FF00FF00\"/><color rgb=\"FF0000FF\"/></colorScale></cfRule>")
  expect_identical(exp, wb$worksheets[[4]]$conditionalFormatting)

})

test_that("extend dataBar tests", {

  wb <- wb_workbook()
  wb$add_worksheet("databar")
  ## Databars
  wb$add_data("databar", -5:5, startCol = 1)
  wb <- wb_add_conditional_formatting(
    wb,
    "databar",
    cols = 1,
    rows = 1:11,
    type = "dataBar"
  ) ## Default colours

  wb$add_data("databar", -5:5, startCol = 3)
  wb <- wb_add_conditional_formatting(
    wb,
    "databar",
    cols = 3,
    rows = 1:11,
    type = "dataBar",
    params = list(
      showValue = FALSE,
      gradient = FALSE
    )
  ) ## Default colours

  wb$add_data("databar", -5:5, startCol = 5)
  wb <- wb_add_conditional_formatting(
    wb,
    "databar",
    cols = 5,
    rows = 1:11,
    type = "dataBar",
    style = c("#a6a6a6"),
    params = list(showValue = FALSE)
  )

  wb$add_data("databar", -5:5, startCol = 7)
  wb <- wb_add_conditional_formatting(
    wb,
    "databar",
    cols = 7,
    rows = 1:11,
    type = "dataBar",
    style = c("red"),
    params = list(
      showValue = TRUE,
      gradient = FALSE
    )
  )

  wb$add_data("databar", -5:5, startCol = 9)
  wb <- wb_add_conditional_formatting(
    wb,
    "databar",
    cols = 9,
    rows = 1:11,
    type = "dataBar",
    style = c("#a6a6a6", "#a6a6a6"),
    params = list(showValue = TRUE, gradient = FALSE)
  )


  # chained test with rule
  wb$
    add_data(x = -5:5, startCol = 11)$
    add_conditional_formatting(
      cols = 11,
      rows = 1:11,
      type = "dataBar",
      rule = c(0, 5),
      style = c("#a6a6a6", "#a6a6a6"),
      params = list(showValue = TRUE, gradient = FALSE)
    )

  exp <- c(`A1:A11` = "<cfRule type=\"dataBar\" priority=\"6\"><dataBar showValue=\"1\"><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40000}</x14:id></ext></extLst></cfRule>",
           `C1:C11` = "<cfRule type=\"dataBar\" priority=\"5\"><dataBar showValue=\"0\"><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40001}</x14:id></ext></extLst></cfRule>",
           `E1:E11` = "<cfRule type=\"dataBar\" priority=\"4\"><dataBar showValue=\"0\"><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFA6A6A6\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40002}</x14:id></ext></extLst></cfRule>",
           `G1:G11` = "<cfRule type=\"dataBar\" priority=\"3\"><dataBar showValue=\"1\"><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFFF0000\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40003}</x14:id></ext></extLst></cfRule>",
           `I1:I11` = "<cfRule type=\"dataBar\" priority=\"2\"><dataBar showValue=\"1\"><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFA6A6A6\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40004}</x14:id></ext></extLst></cfRule>",
           `K1:K11` = "<cfRule type=\"dataBar\" priority=\"1\"><dataBar showValue=\"1\"><cfvo type=\"num\" val=\"0\"/><cfvo type=\"num\" val=\"5\"/><color rgb=\"FFA6A6A6\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40005}</x14:id></ext></extLst></cfRule>"
  )
  got <- wb$worksheets[[1]]$conditionalFormatting
  expect_equal(exp, got)

  exp <- read_xml("<ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">
                  <x14:conditionalFormattings>
                  <x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:cfRule type=\"dataBar\" id=\"{F7189283-14F7-4DE0-9601-54DE9DB40000}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" gradient=\"1\" negativeBarBorderColorSameAsPositive=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:borderColor rgb=\"FF638EC6\"/><x14:negativeFillColor rgb=\"FFFF0000\"/><x14:negativeBorderColor rgb=\"FFFF0000\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A11</xm:sqref></x14:conditionalFormatting>
                  <x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:cfRule type=\"dataBar\" id=\"{F7189283-14F7-4DE0-9601-54DE9DB40001}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" gradient=\"0\" negativeBarBorderColorSameAsPositive=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:borderColor rgb=\"FF638EC6\"/><x14:negativeFillColor rgb=\"FFFF0000\"/><x14:negativeBorderColor rgb=\"FFFF0000\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule><xm:sqref>C1:C11</xm:sqref></x14:conditionalFormatting>
                  <x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:cfRule type=\"dataBar\" id=\"{F7189283-14F7-4DE0-9601-54DE9DB40002}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" gradient=\"1\" negativeBarBorderColorSameAsPositive=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:borderColor rgb=\"FFA6A6A6\"/><x14:negativeFillColor rgb=\"FFFF0000\"/><x14:negativeBorderColor rgb=\"FFFF0000\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule><xm:sqref>E1:E11</xm:sqref></x14:conditionalFormatting>
                  <x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:cfRule type=\"dataBar\" id=\"{F7189283-14F7-4DE0-9601-54DE9DB40003}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" gradient=\"0\" negativeBarBorderColorSameAsPositive=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:borderColor rgb=\"FFFF0000\"/><x14:negativeFillColor rgb=\"FFFF0000\"/><x14:negativeBorderColor rgb=\"FFFF0000\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule><xm:sqref>G1:G11</xm:sqref></x14:conditionalFormatting>
                  <x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:cfRule type=\"dataBar\" id=\"{F7189283-14F7-4DE0-9601-54DE9DB40004}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" gradient=\"0\" negativeBarBorderColorSameAsPositive=\"0\"><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/><x14:borderColor rgb=\"FFA6A6A6\"/><x14:negativeFillColor rgb=\"FFA6A6A6\"/><x14:negativeBorderColor rgb=\"FFA6A6A6\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule><xm:sqref>I1:I11</xm:sqref></x14:conditionalFormatting>
                  <x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:cfRule type=\"dataBar\" id=\"{F7189283-14F7-4DE0-9601-54DE9DB40005}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" border=\"1\" gradient=\"0\" negativeBarBorderColorSameAsPositive=\"0\"><x14:cfvo type=\"num\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"num\"><xm:f>5</xm:f></x14:cfvo><x14:borderColor rgb=\"FFA6A6A6\"/><x14:negativeFillColor rgb=\"FFA6A6A6\"/><x14:negativeBorderColor rgb=\"FFA6A6A6\"/><x14:axisColor rgb=\"FF000000\"/></x14:dataBar></x14:cfRule><xm:sqref>K1:K11</xm:sqref></x14:conditionalFormatting>
                  </x14:conditionalFormattings></ext>", pointer = FALSE)
  got <- wb$worksheets[[1]]$extLst
  expect_equal(exp, got)

})

test_that("wb_conditional_formatting", {
  wb <- wb_workbook()
  wb$add_worksheet("databar")
  ## Databars
  wb$add_data("databar", -5:5, startCol = 1)
  expect_warning(wb <- wb_conditional_formatting(
    wb,
    "databar",
    cols = 1,
    rows = 1:11,
    type = "dataBar"
  ), "deprecated")

  exp <- c(`A1:A11` = "<cfRule type=\"dataBar\" priority=\"1\"><dataBar showValue=\"1\"><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar><extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40000}</x14:id></ext></extLst></cfRule>")
  got <- wb$worksheets[[1]]$conditionalFormatting
  expect_equal(exp, got)
})

test_that("create dxfs style without font family and size", {

  # a workbook with this style loads, but has no highlighting
  exp <- "<dxf><font/></dxf>"
  got <- create_dxfs_style()
  expect_equal(exp, got)

  # most likely what the user expects. change font color and background color
  exp <- "<dxf><font><color rgb=\"FF9C0006\"/></font><fill><patternFill patternType=\"solid\"><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf>"
  got <- create_dxfs_style(
    font_color = wb_colour(hex = "FF9C0006"),
    bgFill = wb_colour(hex = "FFFFC7CE")
  )
  expect_equal(exp, got)

  # the fully fletched old default dxfs style
  exp <- "<dxf><font><color rgb=\"FF9C0006\"/><name val=\"Calibri\"/><sz val=\"11\"/></font><fill><patternFill patternType=\"solid\"><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf>"
  got <- create_dxfs_style(
    font_name = "Calibri",
    font_size = 11,
    font_color = wb_colour(hex = "FF9C0006"),
    bgFill = wb_colour(hex = "FFFFC7CE")
  )
  expect_equal(exp, got)

})
