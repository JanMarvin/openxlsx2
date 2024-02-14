# convert colors from xlsx files to rds file

library(openxlsx2)

# local clone from https://github.com/JanMarvin/openxlsx-data
fls <- dir("../openxlsx-data/colors", pattern = "*.xlsx$", full.names = TRUE)


nms  <- vector("character", length(fls))
clrs <- vector("list", length(fls))

for (i in seq_along(fls)) {

  theme <- wb_load(fls[i])$theme

  clrs[[i]] <- xml_node(theme, "a:theme", "a:themeElements", "a:clrScheme")
  nms[i]    <- xml_attr(clrs[[i]], "a:clrScheme")[[1]][["name"]]

}

names(clrs) <- nms

colors <- clrs

names(colors) %>% dput()

saveRDS(colors, "inst/extdata/colors.rds")
