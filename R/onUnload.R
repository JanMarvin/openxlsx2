.onUnload <- function(libpath) {
  library.dynam.unload("openxlsx2", libpath) #nolint
}
