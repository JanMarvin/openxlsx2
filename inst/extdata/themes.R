# convert themes from xlsx files to rds file

library(openxlsx2)

# local clone from https://github.com/JanMarvin/openxlsx-data
fls <- c(
  dir("../openxlsx-data/styles", full.names = TRUE, pattern = ".xlsx"),
 "../openxlsx-data/loadExample.xlsx"
)

themes <- vector("list", length = length(fls))
nms    <- vector("character", length = length(fls))

for (i in seq_along(fls)) {

  wb <- wb_load(fls[i])

  xml_theme <- wb$theme
  xml_name <- xml_attr(xml_theme, "a:theme")[[1]][["name"]]
  message(xml_name, " ", fls[i])

  themes[[i]] <- stringi::stri_escape_unicode(xml_theme)
  nms[[i]] <- xml_name
}
nms[length(nms)] <- "Old Office Theme"
names(themes) <- nms

themes <- themes[order(names(themes))]

names(themes) %>% dput()

saveRDS(themes, "inst/extdata/themes.rds")
