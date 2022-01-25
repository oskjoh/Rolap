#' Download and install Rolap dependencies
#'
#' This function downloads and installs dependencies (IronPython and 'Microsoft.AnalysisServices.AdomdClient.dll') for the Rolap package.
#' @param ironpython_url Path for downloading Ironpython from the web.
#' @param adomd_client_dll Path for downloading ZIP-file from Nuget containing the necessary Adomd Client DLL file.
#' @param install_dir Installation path of the downloaded dependencies. Defaults to the dependencies folder within the Rolap library directory.
#' @param assume_yes Assume "yes" as answer for prompt to confirm download and installation of dependencies and run non-interactively.
#' @keywords olap MDX
#' @export
#' @examples
#' \dontrun{
#' rolap_download_dependencies()
#'  }

rolap_download_dependencies <- function(provide_credentials = TRUE,
                                        ironpython_url = "https://github.com/IronLanguages/ironpython2/releases/download/ipy-2.7.11/IronPython.2.7.11.zip",
                                        adomd_client_url = "https://www.nuget.org/api/v2/package/Microsoft.AnalysisServices.AdomdClient.retail.amd64",
                                        install_dir =  find.package("Rolap"),
                                        assume_yes = FALSE) {

  # Prepare tempdirs and tempfiles
  tempfile_ironpython <- tempfile()
  tempfile_adomd_client <- tempfile()

  tempdir_ironpython <- tempdir()
  tempdir_adomd_client <- tempdir()

  install_dir_ironpython <- file.path(install_dir, "ironpython")
  install_dir_adomd_client <- file.path(install_dir, "ironpython", "DLLs")

  message("Welcome to Rolap. An R package using IronPython to connect to MSASS Olap Cubes.
        Rolap uses IronPython and 'Microsoft.AnalysisServices.AdomdClient.dll' to connect and query OLAP Cubes.
        This setup script will now download, unzip and install the required dependency files.\n")

  if (assume_yes) {
    continue = TRUE
  } else {
    answer <- readline("Dou you want to continue (Y/N)?")
    continue <- trimws(tolower(answer)) == "y"
  }
  if(!continue) {
    message("Exiting install script.")
  } else {
    message("Downloading and installing IronPython.")

    message(paste("Downloading", ironpython_url))
    utils::download.file(ironpython_url, tempfile_ironpython, method = "curl", extra = c("-L", "--anyauth"))

    message(paste("Unzipping IronPython to:", install_dir_ironpython))
    utils::unzip(tempfile_ironpython, exdir = tempdir_ironpython)
    file.copy(list.files(file.path(tempdir_ironpython, "net45"), full.names = TRUE), install_dir_ironpython,
              overwrite = TRUE, recursive = TRUE)

    message("\nDownloading and installing 'Microsoft.AnalysisServices.AdomdClient.dll'")

    download.file(adomd_client_url, tempfile_adomd_client, method = "curl", extra = c("-L", "--anyauth"))

    message(paste("\nUnzipping 'Microsoft.AnalysisServices.AdomdClient.dll' to:", install_dir_adomd_client))
    unzip(tempfile_adomd_client, files = "lib/net45/Microsoft.AnalysisServices.AdomdClient.dll", exdir = tempdir_adomd_client, overwrite = TRUE)
    quiet <- file.copy(file.path(tempdir_adomd_client, "lib", "net45", "Microsoft.AnalysisServices.AdomdClient.dll"), install_dir_adomd_client)

    if(file.exists(file.path(install_dir, "ironpython","ipy.exe"))
       &
       file.exists(file.path(install_dir, "ironpython", "DLLs", "Microsoft.AnalysisServices.AdomdClient.dll"))
    ) {
      message("\nSetup successfull!")
    } else {
      message("\nSetup finnished without a successfull install!")
    }
  }
}

#' Specify connection parameters to connect to an MS OLAP Cube
#'
#' This function returns a list formated with the correct parameters
#' needed to connect to an MS OLAP Cube using 'Rolap::read_olap'
#' @param userid Username or User ID. Defaults to saved credentials. If saved credentials are not available, Rstudio will ask for credentials.
#' @param password User password. Defaults to saved credentials. If saved credentials are not available, Rstudio will ask for credentials.
#' @param datasource A server address to the OLAP cube.
#' @param catalog The name of the cube.
#' @param fileout storage path for the intermediate csv file.
#' @keywords olap MDX
#' @export
#' @examples
#' \dontrun{
#' con <- olap_connection("user", "pwd", "https://myolapcube.mycompany.com/msolap/", catalog = "Sales")
#' read_olap(
#'  con,
#'  "SELECT
#'     { [Measures].[Pack sales], [Measures].[Value sales] } ON COLUMNS,
#'     { [Year].[2020], [Year].[2021] } ON ROWS
#'   FROM [Sales]
#'   WHERE ( [Products].[Pharmaceuticals].[Oncology] )")
#'  }
olap_connection <- function (userid = NULL, password = NULL, datasource,
                             catalog, fileout = tempfile()) {

  ironpython_path  = file.path(find.package("Rolap"), "ironpython")
  adomd_client_path = file.path(ironpython_path, "DLLs", "Microsoft.AnalysisServices.AdomdClient.dll")

  if(!file.exists(file.path(ironpython_path, "ipy.exe"))) stop(paste0("IronPython not found at: ", ironpython_path,  ". (You can download and installed it by running 'rolap_download_dependencies()'.)"))
  if(!file.exists(file.path(adomd_client_path))) stop(paste(adomd_client_path, "not found! (You can download and installed it by running 'rolap_download_dependencies()'.)"))

  if (is.null(userid) | is.null(password)) {
    if(all(c("Rolap username", "Rolap password") %in% keyring::key_list()$username)) {
      userid <- keyring::key_get("RStudio Keyring Secrets", "Rolap username")
      password <- keyring::key_get("RStudio Keyring Secrets", "Rolap password")
    } else {
      userid <- rstudioapi::askForSecret("Rolap username")
      password <- rstudioapi::askForSecret("Rolap password")
    }
  }
  list(datasource = datasource, catalog = catalog, userid = userid,
       password = password, fileout = fileout, adomd_client_path = adomd_client_path)
}

#' Internal Querying function
#'
#' Common internal function which performs the system calls to the IronPython
#' scripts. Is called by 'rolap::read_olap()' and 'rolap::explore_schema()'.
#'
#' This function returns a tibble of data from an MDX queried OLAP cube.
#' @param con A list of connection parameters, as returned by
#'   'Rolap::olap_connection'.
#' @param query A MDX-query, or level to explore.
#' @param script The name of the underlying python function to call.
#' @param clean_names Clean column names by the use of 'janitor::clean_names'?
#' @keywords olap MDX

query_olap <- function(con,
                       query,
                       call = c("explore_schema", "read_olap"),
                       clean_names = TRUE) {
  # Note the time when we started running
  ts_start <- Sys.time()

  # Get connection info
  coninf <- c(unlist(con), query, call)

  # Prepare python script (paths)
  path2py <- file.path(find.package("Rolap"), "ironpython", "call_from_R.py")
  path2exe <- file.path(find.package("Rolap"), "ironpython", "ipy.exe")

  # Run python script
  args <- paste(path2exe, path2py, paste0("\"", coninf, "\"", collapse = " "))
  call <- suppressWarnings(system(args, intern = TRUE))

  # If system call returns error, show error and stop function.
  if(length(call) != 0) stop(call[length(call)])

  # Before reading file, check that it has actually changed since we called the
  # function (and is not a remainder from a previous run)
  if(file.info(con[["fileout"]])$mtime < ts_start) stop("Internal tempfile not found. Failed to retrive data.")

  # Read the produced csv and return
  .df <- readr::read_delim(con[["fileout"]], "|", col_types = readr::cols())

  if(clean_names) {
    .df <- janitor::clean_names(.df)
  }
  return(.df)
}

#' Query an MS OLAP Cube
#'
#' This function returns a tibble of data from an MDX queried OLAP cube.
#' @param con A list of connection parameters, as returned by 'Rolap::olap_connection'.
#' @param query A MDX-query.
#' @param clean_names Clean column names by the use of 'janitor::clean_names'?
#' @keywords olap MDX
#' @export
#' @examples
#' \dontrun{
#' con <- olap_connection("user", "pwd", "https://myolapcube.mycompany.com/msolap/", catalog = "Sales")
#' read_olap(
#'  con,
#'  "SELECT
#'     { [Measures].[Pack sales], [Measures].[Value sales] } ON COLUMNS,
#'     { [Year].[2020], [Year].[2021] } ON ROWS
#'   FROM [Sales]
#'   WHERE ( [Products].[Pharmaceuticals].[Oncology] )")
#'  }

read_olap <- function(con, query, clean_names = TRUE) {
  query_olap(con, query, "read_olap", clean_names)
}

#' Explore MS OLAP Schema
#'
#' This function returns a tibble of information regarding the available Cubes
#' or the contents of Cubes.
#' @param con A list of connection parameters, as returned by 'Rolap::olap_connection'.
#' @param field The field of information to explore. Read more in details.
#' @param clean_names Clean column names by the use of 'janitor::clean_names'?
#' @details ## Fields
#' The schema of information from the database contains different fields of information.
#' You need to specify which type of information you want to explore. Usually,
#' you should get most of the useful information from the following filelds: 'Cubes', 'Hierarchies',
#' 'Levels' and 'Measures'. All available fields to explore can be found
#' [here](https://docs.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.adomdclient.adomdschemaguid?view=analysisservices-dotnet#fields).
#' @keywords olap MDX
#' @export
#' @examples
#' \dontrun{
#' con <- olap_connection("user", "pwd", "https://myolapcube.mycompany.com/msolap/", catalog = "Sales")
#' explore_schema(con, "Levels")
#'  }

explore_schema <- function(con, query, clean_names = TRUE) {
  query_olap(con, query, "explore_schema", clean_names)
}
