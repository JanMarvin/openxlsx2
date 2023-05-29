download_testfiles <- function() {

  fls <- c(
    "pivot_notes.xlsx",
    "loadExample.xlsx",
    "mtcars_chart.xlsx",
    "tableStyles.xlsx",
    "cloneWorksheetExample.xlsx",
    "cloneEmptyWorksheetExample.xlsx",
    "readTest.xlsx",
    "formula.xlsx",
    "loadPivotTables.xlsx",
    "loadThreadComment.xlsx",
    "inlineStr.xlsx",
    "ColorTabs3.xlsx",
    "namedRegions.xlsx",
    "namedRegions2.xlsx",
    "eurosymbol.xlsx",
    "macro2.xlsm",
    "fichier_complementaire_ccam_descriptive_a_usage_pmsi_2021_v2.xlsx",
    "vml_numbering.xlsx",
    "umlauts.xlsx",
    "unemployment-nrw202208.xlsx",
    "gh_issue_416.xlsm",
    "overwrite_formula.xlsx",
    "connection.xlsx",
    "charts.xlsx",
    "form_control.xlsx",
    "gh_issue_504.xlsx",
    "Single_hyperlink.xlsx",
    "oxlsx2_sheet.xlsx"
  )

  if (dir.exists("testfiles"))  {
    if (all(file.exists(paste0("testfiles/", fls)))) {
      return(TRUE)
    }

    unlink("testfiles", recursive = TRUE)
  }

  dir.create("testfiles")
  testfiles_path <- list.dirs(path = "testfiles")

  for (fl in fls) {
    out <- paste0(testfiles_path, "/", fl)
    url <- paste0("https://github.com/JanMarvin/openxlsx-data/raw/main/", fl)
    download.file(url, destfile = out, quiet = TRUE)
  }

}

download_testfiles()
