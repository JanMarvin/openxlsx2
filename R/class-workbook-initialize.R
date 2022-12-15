wb_initialize_impl <- function(
    self,
    private,
    creator         = NULL,
    title           = NULL,
    subject         = NULL,
    category        = NULL,
    datetimeCreated = Sys.time()
) {
  self$apps <- character()
  self$charts <- list()
  self$isChartSheet <- logical()

  self$connections <- NULL
  self$Content_Types <- genBaseContent_Type()
  self$core <-
    genBaseCore(
      creator = creator,
      title = title,
      subject = subject,
      category = category
    )
  self$comments <- list()
  self$threadComments <- list()


  self$drawings <- list()
  self$drawings_rels <- list()
  # self$drawings_vml <- list()

  self$embeddings <- NULL
  self$externalLinks <- NULL
  self$externalLinksRels <- NULL

  self$headFoot <- NULL

  self$media <- list()
  self$metadata <- NULL

  self$persons <- NULL

  self$pivotTables <- NULL
  self$pivotTables.xml.rels <- NULL
  self$pivotDefinitions <- NULL
  self$pivotRecords <- NULL
  self$pivotDefinitionsRels <- NULL

  self$queryTables <- NULL

  self$slicers <- NULL
  self$slicerCaches <- NULL

  self$sheet_names <- character()
  self$sheetOrder <- integer()

  self$sharedStrings <- list()
  attr(self$sharedStrings, "uniqueCount") <- 0

  self$styles_mgr <- style_mgr$new(self)
  self$styles_mgr$styles <- genBaseStyleSheet()

  self$tables <- NULL
  self$tables.xml.rels <- NULL
  self$theme <- NULL


  self$vbaProject <- NULL
  self$vml <- list()
  self$vml_rels <- list()

  self$creator <-
    creator %||%
    getOption("openxlsx2.creator") %||%
    # USERNAME may only be present for windows
    Sys.getenv("USERNAME", Sys.getenv("USER"))

  assert_class(self$creator,    "character")
  assert_class(title,           "character", or_null = TRUE)
  assert_class(subject,         "character", or_null = TRUE)
  assert_class(category,        "character", or_null = TRUE)
  assert_class(datetimeCreated, "POSIXt")

  stopifnot(
    length(title) <= 1L,
    length(category) <= 1L,
    length(datetimeCreated) == 1L
  )

  self$title           <- title
  self$subject         <- subject
  self$category        <- category
  self$datetimeCreated <- datetimeCreated
  private$generate_base_core()
  private$current_sheet <- 0L
  invisible(self)
}
