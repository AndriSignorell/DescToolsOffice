% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/XLGetRange.R
\name{XLGetRange}
\alias{XLGetRange}
\alias{XLGetWorkbook}
\alias{XLCurrReg}
\alias{XLNamedReg}
\title{Import Data Directly From Excel}
\usage{
XLGetRange(
  file = NULL,
  sheet = NULL,
  range = NULL,
  as.data.frame = TRUE,
  header = FALSE,
  stringsAsFactors = FALSE,
  echo = FALSE,
  na.strings = NULL,
  skip = 0
)

XLGetWorkbook(file, compactareas = TRUE)

XLCurrReg(cell)

XLNamedReg(x)
}
\arguments{
\item{file}{the fully specified path and filename of the workbook. If it is
left as \code{NULL}, the function will look for a running Excel-Application
and use its current sheet. The parameter \code{sheet} will be ignored in
this case.}

\item{sheet}{the name of the sheet containing the range(s) of interest.}

\item{range}{a scalar or a vector with the address(es) of the range(s) to be
returned (characters).  Use "A1"-address mode to specify the ranges, for
example \code{"A1:F10"}. \cr If set to \code{NULL} (which is the default),
the function will look for a selection that contains more than one cell. If
found, the function will use this selection. If there is no selection then
the current region of the selected cell will be used. Use \code{XLCurrReg()}
if the current region of a cell, which is currently not selected, should be
used. Range names can be provided with \code{XLNamedReg("name")}.}

\item{as.data.frame}{logical. Determines if the cellranges should be coerced
into data.frames. Defaults to \code{TRUE}, as this is probably the common
use of this function.}

\item{header}{a logical value indicating whether the range contains the
names of the variables as its first line. Default is \code{FALSE}.
\code{header} is ignored if \code{as.data.frame} has been set to
\code{FALSE}.}

\item{stringsAsFactors}{logical. Should character columns be coerced to
factors? The default is \code{FALSE}, which will return character vectors.}

\item{echo}{logical. If set to \code{TRUE}, the function will print the full
command used, such that it can be copied into the R-script for future use.}

\item{na.strings}{a character vector of strings which are to be interpreted
as \code{NA} values. Blank fields are always considered to be missing
values. Default is \code{NULL}, meaning none.}

\item{skip}{the number of lines of the data file to skip before beginning to
read data.}

\item{compactareas}{logical, defining if areas should be returned by
\code{XLGetWorkbook} as list or as matrix (latter is default).}

\item{cell}{range of the left uppe cell, when current region should be used.}

\item{x}{the name or the index of the XL-name to be used.}
}
\value{
If \code{as.data.frame} is set to \code{TRUE}, a single data.frame
or a list of data.frames will be returned. If set to \code{FALSE} a list of
the cell values in the specified Excel range, resp. a list of lists will be
returned.

\code{XLGetWorkbook()} returns a list of lists of the values in the given
workbook.
}
\description{
The package \code{RDCOMClient} is used to open an Excel workbook and return
the content (value) of one (or several) given range(s) in a specified sheet.
This is helpful, whenever pathologically scattered data on an Excel sheet,
which can't simply be saved as CSV-file, has to be imported in R.\cr\cr
\code{XLGetWorkbook()} does the same for all the sheets in an Excel
workbook.
}
\details{
The result consists of a list of lists, if \code{as.data.frame} is set to
\code{FALSE}. Be then prepared to encounter \code{NULL} values. Those will
prevent from easily being able to coerce the square data structure to a
data.frame.

The following code will replace the \code{NULL} values by \code{NA} and
coerce the data to a data.frame. \preformatted{ # get the range D1:J69 from
an excel file xlrng <- XLGetRange(file="myfile.xlsx", sheet="Tabelle1",
range="D1:J69", as.data.frame=FALSE)

# replace NULL values by NA xlrng[unlist(lapply(xlrng, is.null))] <- NA

# coerce the square data structure to a data.frame d.lka <-
data.frame(lapply(data.frame(xlrng), unlist)) } This of course can be
avoided by setting \code{as.data.frame} = \code{TRUE}.

The function will return dates as integers, because MS-Excel stores them
internally as integers. Such a date can subsequently be converted with the
(unusual) origin of \code{as.Date(myDate, origin="1899-12-30")}. See also
\code{\link{XLDateToPOSIXct}}, which does the job. The conversion can
directly be performed by \code{XLGetRange()} if \code{datecols} is used and
contains the date columns in the sheet data.
}
\examples{

\dontrun{ # Windows-specific example

XLGetRange(file="C:/My Documents/data.xls",
           sheet="Sheet1",
           range=c("A2:B5","M6:X23","C4:D40"))


# if the current region has to be read (incl. a header), place the cursor in the interesting region
# and run:
d.set <- XLGetRange(header=TRUE)

# Get XL nameslist
nm <- xl$ActiveWorkbook()$names()

lst <- list()
for(i in 1:nm$count())
  lst[[i]] <- c(name=nm[[i]]$name(), 
                address=nm[[i]]$refersToRange()$Address())
  
# the defined names
as.data.frame(do.call(rbind, lst), stringsAsFactors = FALSE)
}

}
\seealso{
\code{\link{GetNewXL}}, \code{\link{GetCurrXL}},
\code{\link{XLView}}
}
\author{
Andri Signorell \href{mailto:andri@signorell.net}{andri@signorell.net}
}
\keyword{manip}
