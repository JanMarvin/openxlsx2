
test_that("Worksheet Class works", {
  expect_null(assert_worksheet(wb_worksheet()))
})

test_that("test data validation list and sparklines", {

  set.seed(123) # sparklines has a random uri string
  options("openxlsx2_seed" = NULL)

  s1 <- create_sparklines("Sheet 1", "A3:K3", "L3")
  s2 <- create_sparklines("Sheet 1", "A4:K4", "L4")

  wb <- wb_workbook()$
    add_worksheet()$add_data(x = iris[1:30,])$
    add_worksheet()$add_data(sheet = 2, x = sample(iris$Sepal.Length, 10))$
    add_data_validation(sheet = 1, col = 1, rows = 2:11, type = "list", value = '"O1,O2"')$
    add_sparklines(sheet = 1, sparklines = s1)$
    add_data_validation(sheet = 1, col = 1, rows = 12:21, type = "list", value = '"O2,O3"')$
    add_sparklines(sheet = 1, sparklines = s2)

  exp <- c(
    "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\"><x14:dataValidations xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" count=\"2\"><x14:dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\"><x14:formula1><xm:f>\"O1,O2\"</xm:f></x14:formula1><xm:sqref>A2:A11</xm:sqref></x14:dataValidation><x14:dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\"><x14:formula1><xm:f>\"O2,O3\"</xm:f></x14:formula1><xm:sqref>A12:A21</xm:sqref></x14:dataValidation></x14:dataValidations></ext>",
    "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:sparklineGroup displayEmptyCellsAs=\"gap\" xr2:uid=\"{6F57B887-24F1-C14A-942C-ASEVX1JWJGYG}\"><x14:colorSeries rgb=\"FF376092\"/><x14:colorNegative rgb=\"FFD00000\"/><x14:colorAxis rgb=\"FFD00000\"/><x14:colorMarkers rgb=\"FFD00000\"/><x14:colorFirst rgb=\"FFD00000\"/><x14:colorLast rgb=\"FFD00000\"/><x14:colorHigh rgb=\"FFD00000\"/><x14:colorLow rgb=\"FFD00000\"/><x14:sparklines><x14:sparkline><xm:f>'Sheet 1'!A3:K3</xm:f><xm:sqref>L3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup><x14:sparklineGroup displayEmptyCellsAs=\"gap\" xr2:uid=\"{6F57B887-24F1-C14A-942C-9DKW7WYNM276}\"><x14:colorSeries rgb=\"FF376092\"/><x14:colorNegative rgb=\"FFD00000\"/><x14:colorAxis rgb=\"FFD00000\"/><x14:colorMarkers rgb=\"FFD00000\"/><x14:colorFirst rgb=\"FFD00000\"/><x14:colorLast rgb=\"FFD00000\"/><x14:colorHigh rgb=\"FFD00000\"/><x14:colorLow rgb=\"FFD00000\"/><x14:sparklines><x14:sparkline><xm:f>'Sheet 1'!A4:K4</xm:f><xm:sqref>L4</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"
  )
  got <- wb$worksheets[[1]]$extLst
  expect_equal(exp, got)

})
