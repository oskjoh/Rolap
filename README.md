# Rolap - An open source R OLAP connection package.

## Introduction

Rolap is a package that lets R connect to Microsoft Analysis Services OLAP
qubes and query them through MDX-queries. It is similar to the 'olapR' package
that is available as part of the 'Microsoft R open'-client but can be used with
the regular R distribution.

The package is using [IronPython](https://ironpython.net/) and 
'Microsoft.AnalysisServices.AdomdClient.dll' DLL library behind the scenes.
An IronPython-script is called by R, which queries the database and returns a
temporary CSV-file which is then read into R.

Because this package relies on IronPython and the ADOMD-Client DLL, it is only
working on Windows systems (for the moment).

## Installation

### 1. Install the package

Install the package using 
    ```
    devtools::install_github("oskjoh/Rolap")
    ```
.

### 2. Download and install dependencies

The package needs access to an IronPython-distribution and the DLL-file
'Microsoft.AnalysisServices.AdomdClient.dll'. These dependencies can be
installed using one of the following two methods:

#### Method 1
    
After installing the package in R. Run the provided installation script
    ```
    Rolap::rolap_download_dependencies()
    ```
to download and install the required dependencies.

#### Method 2

Download the dependencies yourself and place the required files in the
```Rolap/ironpython``` folder under your R library path. (The correct folder
can usually be found by running ```find.package("Rolap")```.)

1. Download the [zip version of IronPython 2.7](https://github.com/IronLanguages/ironpython2/releases/download/ipy-2.7.11/IronPython.2.7.11.zip).
Unzip the downloaded file and copy the contents of the ```net45``` folder to the
```Rolap/ironpython``` folder.

2. Locate/Download ```Microsoft.AnalysisServices.AdomdClient.dll``` and place it in the
```Rolap/ironpython/DLLs``` folder. The 'Rolap::rolap_download_dependencies()'-script does this
by downloading the Nuget-package-file from [here](https://www.nuget.org/packages/Microsoft.AnalysisServices.AdomdClient.retail.amd64/) and obtains the required DLL file from the Nuget-package (which is really just a zip-archive. The file is located in the archive directory ```lib/net45```

## Usage

Now you should be ready to make your first query! 

If no username or password
is provided. The package will look for saved credentials using keyring. If 
no such credentials are available. Rolap will ask for your username and password
and allow you to save them for later.

```{r}
library(Rolap)

con <- olap_connection("user", "pwd", "https://myolapcube.mycompany.com/msolap/", catalog = "Sales")
read_olap(
con,
"SELECT
{ [Measures].[Pack sales], [Measures].[Value sales] } ON COLUMNS,
{ [Year].[2020], [Year].[2021] } ON ROWS
FROM [Sales]
WHERE ( [Products].[Pharmaceuticals].[Oncology] )")

```



