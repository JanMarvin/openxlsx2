.onUnload <- function(libpath) {
  gc() # trigger finalizers
  library.dynam.unload("openxlsx2", libpath)
}
