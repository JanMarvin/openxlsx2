library(magick)

fls <- dir(path = "inst/img", pattern = ".png", full.names = TRUE)

for (fl in fls) {
  message(fl)

  img <- magick::image_read(fl)

  path <- paste0("vignettes/img/", basename(fl))

  defines <- c("png:compression-filter" = "1",
               "png:compression-level" = "9")
  image_set_defines(img, defines)

  # which option reduces the file size ...
  magick::image_write(image = img, path = path)

  # compression
  in_s <- file.size(fl)
  out_s <- file.size(path)
  print(out_s/in_s)

  if (out_s > in_s) {
    file.copy(fl, path, overwrite = TRUE)
  }

}
