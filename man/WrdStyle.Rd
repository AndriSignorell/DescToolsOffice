\name{WrdStyle}
\alias{WrdStyle}
\alias{WrdStyle<-}
\title{Get or Set the Style in Word
}
\description{\code{WrdStyle} can be used to get and set the style in Word for the text to be inserted. \code{WrdStyle} returns the style at the current cursor position.
}
\usage{
WrdStyle(wrd = DescToolsOfficeOptions("lastWord"))
WrdStyle(wrd) <- value
}
\arguments{
  \item{value}{the name of the style to be used to the output. This should be defined an existing name.
}
  \item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.}

}
\value{
character, name of the style

}
\author{
Andri Signorell <andri@signorell.net>
}

\seealso{\code{\link{ToWrd}}, \code{\link{WrdPlot}}, \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
}
\examples{
\dontrun{ # Windows-specific example

wrd <- GetNewWrd()
# the current stlye
WrdStyle(wrd)
}}
\keyword{ print }
\concept{MS-Office}
