\name{WrdSaveAs}
\alias{WrdSaveAs}
\alias{WrdOpenFile}
\title{Open and Save Word Documents
}
\description{Open and save MS-Word documents.
}
\usage{
WrdOpenFile(fn, wrd = DescToolsOfficeOptions("lastWord"))
WrdSaveAs(fn, fileformat = "docx", wrd = DescToolsOfficeOptions("lastWord"))
}
\arguments{
  \item{fn}{filename and -path for the document.
}
  \item{fileformat}{file format, one out of \code{"doc"}, \code{"htm"}, \code{"pdf"}.
}
  \item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.}

}
\value{nothing returned
}
\author{Andri Signorell <andri@signorell.net>
}

\seealso{\code{\link{GetNewWrd}()}
}
\examples{
\dontrun{
#   Windows-specific example
wrd <- GetNewWrd()
WrdCaption("A Report")
WrdSaveAs(fn="report", fileformat="htm")
}}
\keyword{print}
\concept{MS-Office}
