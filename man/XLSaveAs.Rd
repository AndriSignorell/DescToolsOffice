\name{XLSaveAs}
\alias{XLSaveAs}
\title{Save Excel File
}
\description{Save the current workbook under the given name and format.
}
\usage{
XLSaveAs(fn, file_format = xlConst$XlFileFormat$xlWorkbookNormal, 
         xl = DescToolsOfficeOptions("lastXL"))
}
\arguments{
  \item{fn}{the filename
}
  \item{file_format}{the file format using the xl constant.
}
\item{xl}{the pointer to a MS-Excel instance. An new instance can be created with \code{GetNewXL()}, returning the appropriate handle. A handle to an already running instance is returned by \code{GetCurrXL()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastXL")}.}
}
\value{returns \code{TRUE} if the save operation has been successful
}
\author{Andri Signorell <andri@signorell.net>
}
\seealso{\code{\link{XLView}}
}
\examples{
\dontrun{# Windows-specific example
XLView(d.diamonds)
XLSaveAs("Diamonds")
xl$quit()
}}
\keyword{ manip }
\concept{MS-Office}