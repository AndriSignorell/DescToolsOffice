\name{WrdPageBreak}
\alias{WrdPageBreak}
\title{Insert a Page Break
}
\description{Insert a page break in a MS-Word (R) document at the position of the cursor.
}
\usage{
WrdPageBreak(wrd = DescToolsOfficeOptions("lastWord"))
}
\arguments{
 \item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.}
}

\author{Andri Signorell <andri@signorell.net>
}


\seealso{
\code{\link{WrdFont}}, \code{\link{WrdPlot}}, \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
}
\examples{
\dontrun{ # Windows-specific example
wrd <- GetNewWrd()
WrdText("This is text on page 1.\n\n")
WrdPageBreak()
WrdText("This is text on another page.\n\n")
}}
\keyword{ print }
\concept{MS-Office}
