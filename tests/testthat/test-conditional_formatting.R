test_that("conditional formatting", {

  wb <- wb_workbook()
  wb$add_worksheet("cellIs")

  negStyle <- create_dxfs_style(font_color = c(rgb = "FF9C0006"), bgFill = c(rgb = "FFFFC7CE"))
  posStyle <- create_dxfs_style(font_color = c(rgb = "FF006100"), bgFill = c(rgb = "FFC6EFCE"))

  wb$styles_mgr$styles$dxfs <- c(wb$styles_mgr$styles$dxfs,
                                 c(negStyle, posStyle)
  )

  exp <- c(
    "<dxf><font><color rgb=\"FF9C0006\"/><name val=\"Calibri\"/><sz val=\"11\"/></font><fill><patternFill patternType=\"solid\"><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf>",
    "<dxf><font><color rgb=\"FF006100\"/><name val=\"Calibri\"/><sz val=\"11\"/></font><fill><patternFill patternType=\"solid\"><bgColor rgb=\"FFC6EFCE\"/></patternFill></fill></dxf>"
  )

  expect_equal(exp, wb$styles_mgr$styles$dxfs)

  set.seed(123)

  ## rule applies to all each cell in range
  wb$add_data("cellIs", -5:5)
  wb$add_data("cellIs", LETTERS[1:11], startCol = 2)
  wb_conditional_formatting(wb, "cellIs",
                        cols = 1,
                        rows = 1:11, rule = "!=0", style = negStyle
  )
  wb_conditional_formatting(wb, "cellIs",
                        cols = 1,
                        rows = 1:11, rule = "==0", style = posStyle
  )

  exp <- c(
    `A1:A11` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"2\"><formula>A1&lt;&gt;0</formula></cfRule>",
    `A1:A11` = "<cfRule type=\"expression\" dxfId=\"1\" priority=\"1\"><formula>A1=0</formula></cfRule>"
  )
  expect_equal(exp, wb$worksheets[[1]]$conditionalFormatting)


  wb$add_worksheet("Moving Row")
  ## highlight row dependent on first cell in row
  wb$add_data("Moving Row", -5:5)
  wb$add_data("Moving Row", LETTERS[1:11], startCol = 2)
  wb_conditional_formatting(wb, "Moving Row",
                        cols = 1:2,
                        rows = 1:11, rule = "$A1<0", style = negStyle
  )
  wb_conditional_formatting(wb, "Moving Row",
                        cols = 1:2,
                        rows = 1:11, rule = "$A1>0", style = posStyle
  )

  exp <- c(
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"2\"><formula>$A1&lt;0</formula></cfRule>",
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"1\" priority=\"1\"><formula>$A1&gt;0</formula></cfRule>"
  )
  expect_equal(exp, wb$worksheets[[2]]$conditionalFormatting)

  wb$add_worksheet("Moving Col")
  ## highlight column dependent on first cell in column
  wb$add_data("Moving Col", -5:5)
  wb$add_data("Moving Col", LETTERS[1:11], startCol = 2)
  wb_conditional_formatting(wb, "Moving Col",
                        cols = 1:2,
                        rows = 1:11, rule = "A$1<0", style = negStyle
  )
  wb_conditional_formatting(wb, "Moving Col",
                        cols = 1:2,
                        rows = 1:11, rule = "A$1>0", style = posStyle
  )

  exp <- c(
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"2\"><formula>A$1&lt;0</formula></cfRule>",
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"1\" priority=\"1\"><formula>A$1&gt;0</formula></cfRule>"
  )
  expect_equal(exp, wb$worksheets[[3]]$conditionalFormatting)


  wb$add_worksheet("Dependent on")
  ## highlight entire range cols X rows dependent only on cell A1
  wb$add_data("Dependent on", -5:5)
  wb$add_data("Dependent on", LETTERS[1:11], startCol = 2)
  wb_conditional_formatting(wb, "Dependent on",
                        cols = 1:2,
                        rows = 1:11, rule = "$A$1<0", style = negStyle
  )
  wb_conditional_formatting(wb, "Dependent on",
                        cols = 1:2,
                        rows = 1:11, rule = "$A$1>0", style = posStyle
  )

  exp <- c(
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"2\"><formula>$A$1&lt;0</formula></cfRule>",
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"1\" priority=\"1\"><formula>$A$1&gt;0</formula></cfRule>"
  )
  expect_equal(exp, wb$worksheets[[4]]$conditionalFormatting)


  ## highlight cells in column 1 based on value in column 2
  wb$add_data("Dependent on", data.frame(x = 1:10, y = runif(10)), startRow = 15)
  wb_conditional_formatting(wb, "Dependent on",
                        cols = 1,
                        rows = 16:25, rule = "B16<0.5", style = negStyle
  )
  wb_conditional_formatting(wb, "Dependent on",
                        cols = 1,
                        rows = 16:25, rule = "B16>=0.5", style = posStyle
  )

  exp <- c(
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"4\"><formula>$A$1&lt;0</formula></cfRule>",
    `A1:B11` = "<cfRule type=\"expression\" dxfId=\"1\" priority=\"3\"><formula>$A$1&gt;0</formula></cfRule>",
    `A16:A25` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"2\"><formula>B16&lt;0.5</formula></cfRule>",
    `A16:A25` = "<cfRule type=\"expression\" dxfId=\"1\" priority=\"1\"><formula>B16&gt;=0.5</formula></cfRule>"
  )
  expect_equal(exp, wb$worksheets[[4]]$conditionalFormatting)

  wb$add_worksheet("Duplicates")
  ## highlight duplicates using default style
  wb$add_data("Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
  wb_conditional_formatting(wb, "Duplicates", cols = 1, rows = 1:10, type = "duplicates")

  exp <- c(`A1:A10` = "<cfRule type=\"duplicateValues\" dxfId=\"0\" priority=\"1\"/>")
  expect_equal(exp, wb$worksheets[[5]]$conditionalFormatting)

  wb$add_worksheet("containsText")
  ## cells containing text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("containsText", sapply(1:10, fn))
  wb_conditional_formatting(wb, "containsText", cols = 1, rows = 1:10, type = "contains", rule = "A")

  # TODO remove identing from xml
  exp <- c(`A1:A10` = "<cfRule type=\"containsText\" dxfId=\"0\" priority=\"1\" operator=\"containsText\" text=\"A\">\n                        \t<formula>NOT(ISERROR(SEARCH(\"A\", A1)))</formula>\n                       </cfRule>")
  expect_equal(exp, wb$worksheets[[6]]$conditionalFormatting)

  wb$add_worksheet("notcontainsText")
  ## cells not containing text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("containsText", sapply(1:10, fn))
  wb_conditional_formatting(wb, "notcontainsText", cols = 1,
                        rows = 1:10, type = "notcontains", rule = "A")

  exp <- c(`A1:A10` = "<cfRule type=\"notContainsText\" dxfId=\"0\" priority=\"1\" operator=\"notContains\" text=\"A\">\n                        \t<formula>ISERROR(SEARCH(\"A\", A1))</formula>\n                       </cfRule>")
  expect_equal(exp, wb$worksheets[[7]]$conditionalFormatting)

  wb$add_worksheet("beginsWith")
  ## cells begins with text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("beginsWith", sapply(1:100, fn))
  wb_conditional_formatting(wb, "beginsWith", cols = 1, rows = 1:100, type = "beginsWith", rule = "A")

  exp <- c(`A1:A100` = "<cfRule type=\"beginsWith\" dxfId=\"0\" priority=\"1\" operator=\"beginsWith\" text=\"A\">\n                        \t<formula>LEFT(A1,LEN(\"A\"))=\"A\"</formula>\n                       </cfRule>")
  expect_equal(exp, wb$worksheets[[8]]$conditionalFormatting)

  wb$add_worksheet("endsWith")
  ## cells ends with text
  fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
  wb$add_data("endsWith", sapply(1:100, fn))
  wb_conditional_formatting(wb, "endsWith", cols = 1, rows = 1:100, type = "endsWith", rule = "A")

  exp <- c(`A1:A100` = "<cfRule type=\"endsWith\" dxfId=\"0\" priority=\"1\" operator=\"endsWith\" text=\"A\">\n                        \t<formula>RIGHT(A1,LEN(\"A\"))=\"A\"</formula>\n                       </cfRule>")
  expect_equal(exp, wb$worksheets[[9]]$conditionalFormatting)

  wb$add_worksheet("colourScale", zoom = 30)
  ## colourscale colours cells based on cell value
  df <- read_xlsx(system.file("extdata", "readTest.xlsx", package = "openxlsx2"), sheet = 4)
  wb$add_data("colourScale", df, colNames = FALSE) ## write data.frame
  ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
  ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
  wb_conditional_formatting(wb, "colourScale",
                        cols = seq_along(df), rows = seq_len(nrow(df)),
                        style = c("black", "white"),
                        rule = c(0, 255),
                        type = "colourScale"
  )
  wb$set_col_widths("colourScale", cols = seq_along(df), widths = 1.07)
  wb <- wb_set_row_heights(wb, "colourScale", rows = seq_len(nrow(df)), heights = 7.5)

  exp <- c(`A1:E5` = "<cfRule type=\"colorScale\" priority=\"1\"><colorScale>\n                            <cfvo type=\"num\" val=\"0\"/><cfvo type=\"num\" val=\"255\"/>\n                            <color rgb=\"FF000000\"/><color rgb=\"FFFFFFFF\"/>\n                           </colorScale></cfRule>")
  expect_equal(exp, wb$worksheets[[10]]$conditionalFormatting)

  wb$add_worksheet("databar")
  ## Databars
  wb$add_data("databar", -5:5)
  wb_conditional_formatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Default colours

  exp <- c(`A1:A11` = "<cfRule type=\"dataBar\" priority=\"1\"><dataBar showValue=\"1\">\n                          <cfvo type=\"min\"/><cfvo type=\"max\"/>\n                          <color rgb=\"FF638EC6\"/>\n                          </dataBar>\n                          <extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40000}</x14:id></ext>\n                        </extLst></cfRule>")
  expect_equal(exp, wb$worksheets[[11]]$conditionalFormatting)

  wb$add_worksheet("between")
  ## Between
  # Highlight cells in interval [-2, 2]
  wb$add_data("between", -5:5)
  wb_conditional_formatting(wb, "between", cols = 1, rows = 1:11, type = "between", rule = c(-2, 2))

  exp <- c(`A1:A11` = "<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"between\"><formula>-2</formula><formula>2</formula></cfRule>")
  expect_equal(exp, wb$worksheets[[12]]$conditionalFormatting)

  wb$add_worksheet("topN")
  ## Top N
  wb$add_data("topN", data.frame(x = 1:10, y = rnorm(10)))
  # Highlight top 5 values in column x
  wb_conditional_formatting(wb, "topN", cols = 1, rows = 2:11,
                        style = posStyle, type = "topN", rank = 5)#'
  # Highlight top 20 percentage in column y
  wb_conditional_formatting(wb, "topN", cols = 2, rows = 2:11,
                        style = posStyle, type = "topN", rank = 20, percent = TRUE)

  exp <- c(
    `A2:A11` = "<cfRule type=\"top10\" dxfId=\"1\" priority=\"2\" rank=\"5\" percent=\"NULL\"></cfRule>",
    `B2:B11` = "<cfRule type=\"top10\" dxfId=\"1\" priority=\"1\" rank=\"20\" percent=\"1\"></cfRule>"
  )
  expect_equal(exp, wb$worksheets[[13]]$conditionalFormatting)


  wb$add_worksheet("bottomN")
  ## Bottom N
  wb$add_data("bottomN", data.frame(x = 1:10, y = rnorm(10)))
  # Highlight bottom 5 values in column x
  wb_conditional_formatting(wb, "bottomN", cols = 1, rows = 2:11,
                        style = negStyle, type = "topN", rank = 5)
  # Highlight bottom 20 percentage in column y
  wb_conditional_formatting(wb, "bottomN", cols = 2, rows = 2:11,
                        style = negStyle, type = "topN", rank = 20, percent = TRUE)

  exp <- c(
    `A2:A11` = "<cfRule type=\"top10\" dxfId=\"0\" priority=\"2\" rank=\"5\" percent=\"NULL\"></cfRule>",
    `B2:B11` = "<cfRule type=\"top10\" dxfId=\"0\" priority=\"1\" rank=\"20\" percent=\"1\"></cfRule>"
  )
  expect_equal(exp, wb$worksheets[[14]]$conditionalFormatting)


  wb$add_worksheet("logical operators")
  ## Logical Operators
  # You can use Excels logical Operators
  wb$add_data("logical operators", 1:10)
  wb_conditional_formatting(wb, "logical operators",
                        cols = 1, rows = 1:10,
                        rule = "OR($A1=1,$A1=3,$A1=5,$A1=7)"
  )

  exp <- c(`A1:A10` = "<cfRule type=\"expression\" dxfId=\"0\" priority=\"1\"><formula>OR($A1=1,$A1=3,$A1=5,$A1=7)</formula></cfRule>")
  expect_equal(exp, wb$worksheets[[15]]$conditionalFormatting)

  # test saving
  tmp <- temp_xlsx()
  wb_save(wb, tmp)

  expect_silent(wb1 <- wb_load(tmp))

  ## Test fails because of reading and writing
  # # all.equal(wb, wb1)
  # for (sheet in seq_along(wb$sheet_names))
  #   expect_equal(
  #     read_xml(wb$worksheets[[sheet]]$conditionalFormatting, pointer = FALSE),
  #     read_xml(wb1$worksheets[[sheet]]$conditionalFormatting, pointer = FALSE)
  #   )

})
