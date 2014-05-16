excelIO
=======

an R package for I/O of Microsoft Excel files

Installation
------------
Ensure that the perl module Spreadsheet::WriteExcel is installed.
    perl -MCPAN -e 'install "Spreadsheet::WriteExcel"'

Install the R package

    install.packages("devtools")
    library(devtools)
    install_github("excelIO", "drmjc")

Key functions
-----
`read.xls`: imports a worksheet from an Excel (XLS/XLSX) file.

`write.xls`: writes `matrix` or `data.frame` objects, or a `list` of `matrix` or `data.frame`s into worksheet(s) in an XLS file. It also formats the header row, and allows more control over how the `row.names` are written.

`writeLines.xls` writes vector's to excel files as 1-column tables, or as a 2-column table if the `names=TRUE` is specified.
