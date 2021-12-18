.onLoad <- function(libname, pkgname) {

  package_dir = find.package("Rolap")

  if(!(file.exists(file.path(package_dir, "ironpython","ipy.exe"))
       &
       file.exists(file.path(package_dir, "ironpython", "DLLs", "Microsoft.AnalysisServices.AdomdClient.dll"))
  )) {
    message("Cannot find the required dependencies (IronPython/'Microsoft.AnalysisServices.AdomdClient.dll') needed for Rolap to work. You can download and install them by running 'Rolap::rolap_download_dependencies()'.")
  }
}
