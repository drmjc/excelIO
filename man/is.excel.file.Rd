\name{is.excel.file}
\alias{is.excel.file}
\alias{is.xls.file}
\alias{is.xlsx.file}
\title{is the file a Microsoft Excel XLS or XLSX file?}
\usage{
  is.excel.file(path)

  is.xls.file(path)

  is.xlsx.file(path)
}
\arguments{
  \item{path}{a character vector: the path(s) to file(s)
  which must exist}
}
\value{
  A \code{logical vector} of \code{length(path)}, where the
  values are \code{TRUE} if the path is an XLS, XLSX, or
  both, \code{FALSE} otherwise, for
  \code{\link{is.xls.file}}, \code{\link{is.xlsx.file}}, or
  \code{\link{is.excel.file}}, respectively.
}
\description{
  Rather than just checking the file extension, this
  function checks whether the contents of the file are a
  Microsoft Excel file, using the \code{GNU file} program,
  thus this will not work on Windows. \code{file} works by
  checking the file signature to see what file type it is.
  Note that XLS files can have different file signatures
  depending where they are saved from/created.
}
\examples{
\dontrun{
# df <- data.frame(A=1:10, B=letters[1:10], C=LETTERS[1:10], D=TRUE, E=FALSE)
require(datasets)
write.table(iris, "tmp.tsv", quote=FALSE, sep="\\t")
write.csv(iris, "tmp.csv")
require(dataframes2xls)
write.xls(iris, "tmp.xls")
require(xlsx)
write.xlsx(iris, "tmp.xlsx")

is.excel.file("./tmp.xls")
is.excel.file("./tmp.xlsx")
is.excel.file("./tmp.tsv")
is.excel.file("./tmp.csv")
is.excel.file(c("./tmp.xls", "./tmp.xlsx", "./tmp.tsv", "./tmp.csv"))
}
}
\author{
  Mark Cowley, 2012-02-01
}
\seealso{
  \code{\link{is.xls.file}} \code{\link{is.xlsx.file}}
}

