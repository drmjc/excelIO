\name{write.xls}
\alias{write.xls}
\title{Write to Microsoft Excel XLS file}
\usage{
  write.xls(x, xls, verbose = FALSE, row.names = FALSE,
    col.names = TRUE, ...)
}
\arguments{
  \item{x}{either a data.frame, or a list of data.frames.
  if a list of data.frames, then each data.frame will
  become a worksheet in the resulting workbook.}

  \item{xls}{the xls file to be created.}

  \item{verbose}{TRUE/FALSE}

  \item{row.names}{DIFFERENT TO \code{\link{write.table}}.
  Can be either \code{TRUE}, \code{FALSE}, \code{NULL},
  \code{NA}, or a \code{character(1)}. if \code{TRUE} then
  row.names will be written out, with a blank column
  header, if \code{FALSE}, then no row names will be
  written out, if \code{NULL} or \code{NA}, then rownames
  will be written out with no column name (just like
  \code{write.table(row.names=TRUE)} would do) if
  \code{row.names} is a character vector of length 1, then
  write the \code{row.names} out, and use use the supplied
  word as the column name for the rownames else error}

  \item{col.names}{logical: If \code{TRUE} then write out
  the column names.}

  \item{\dots}{additional arguments passed to
  \code{\link{write.delim}}}
}
\value{
  none. writes an excel 'xls' file.
}
\description{
  Write a \code{matrix} or \code{data.frame} to a genuine
  Microsoft Excel file. The first row is bold, italics,
  dashed underline. Columns are centred.
}
\details{
  Uses a perl module called \code{Spreadsheet::WriteExcel},
  by John McNamara, a perl program called \code{tab2xls},
  based on a program of the same name also by JM. This R
  code borrows the structure from \code{gdata}'s excellent
  \code{\link[gdata]{read.xls}} function.
}
\note{
  Users must install the Spreadsheet::WriteExcel Perl
  module.\cr For OSX users, this involves:\cr sudo perl -e
  shell -MCPAN\cr install Spreadsheet::WriteExcel\cr - or
  -\cr sudo cpan\cr ### accept the defaults, and set up
  which country you are in.\cr install CPAN\cr reload
  cpan\cr install Spreadsheet::WriteExcel\cr
}
\author{
  Mark Cowley, 2009-01-20
}
\seealso{
  \code{\link{write.table}} \code{\link{read.delim}}
  \code{\link[gdata]{read.xls}}
}

