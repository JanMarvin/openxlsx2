# requires pngcrush

fls <- dir(path = "inst/img", pattern = ".png", full.names = TRUE)

for (fl in fls) {
  message(fl)
  path <- paste0("vignettes/img/", basename(fl))

  system(paste("pngcrush", fl, path))
}
