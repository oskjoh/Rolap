% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/rolap.R
\name{rolap_download_dependencies}
\alias{rolap_download_dependencies}
\title{Download and install Rolap dependencies}
\usage{
rolap_download_dependencies(
  provide_credentials = TRUE,
 
    ironpython_url = "https://github.com/IronLanguages/ironpython2/releases/download/ipy-2.7.11/IronPython.2.7.11.zip",
 
    adomd_client_url = "https://www.nuget.org/api/v2/package/Microsoft.AnalysisServices.AdomdClient.retail.amd64",
  install_dir = find.package("Rolap"),
  assume_yes = FALSE
)
}
\arguments{
\item{ironpython_url}{Path for downloading Ironpython from the web.}

\item{install_dir}{Installation path of the downloaded dependencies. Defaults to the dependencies folder within the Rolap library directory.}

\item{assume_yes}{Assume "yes" as answer for prompt to confirm download and installation of dependencies and run non-interactively.}

\item{adomd_client_dll}{Path for downloading ZIP-file from Nuget containing the necessary Adomd Client DLL file.}
}
\description{
This function downloads and installs dependencies (IronPython and 'Microsoft.AnalysisServices.AdomdClient.dll') for the Rolap package.
}
\examples{
\dontrun{
rolap_download_dependencies()
 }
}
\keyword{MDX}
\keyword{olap}
