download_testfiles <- function() {

  fls <- c(
    "charts.xlsx",
    "cloneEmptyWorksheetExample.xlsx",
    "cloneWorksheetExample.xlsx",
    "ColorTabs3.xlsx",
    "connection.xlsx",
    "eurosymbol.xlsx",
    "fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx",
    "form_control.xlsx",
    "formula.xlsx",
    "gh_issue_416.xlsm",
    "gh_issue_504.xlsx",
    "inline_str.xlsx",
    "inlineStr.xlsx",
    "loadExample.xlsx",
    "loadPivotTables.xlsx",
    "loadThreadComment.xlsx",
    "macro2.xlsm",
    "mtcars_chart.xlsx",
    "namedRegions.xlsx",
    "namedRegions2.xlsx",
    "overwrite_formula.xlsx",
    "oxlsx2_sheet.xlsx",
    "pivot_notes.xlsx",
    "readTest.xlsx",
    "Single_hyperlink.xlsx",
    "tableStyles.xlsx",
    "umlauts.xlsx",
    "unemployment-nrw202208.xlsx",
    "update_test.xlsx",
    "vml_numbering.xlsx"
  )

  if (dir.exists("testfiles"))  {
    if (all(file.exists(paste0("testfiles/", fls)))) {
      return(TRUE)
    }

    unlink("testfiles", recursive = TRUE)
  }

  dir.create("testfiles")
  testfiles_path <- list.dirs(path = "testfiles")

  # relies on libcurl and was optional in R < 4.2.0 on Windows
  out <- paste0(testfiles_path, "/", fls)
  url <- paste0("https://github.com/JanMarvin/openxlsx-data/raw/main/", fls)
  try({download.file(url, destfile = out, quiet = TRUE, method = "libcurl")})

}

download_testfiles()
