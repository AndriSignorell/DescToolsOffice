\name{WrdCaption}
\alias{WrdCaption}

\title{Insert Caption to Word}
\description{Insert a caption in a given level to a Word document. The caption is inserted at the current cursor position.
}
\usage{
WrdCaption(x, index = 1, wrd = DescToolsOfficeOptions("lastWord"))
}
\arguments{
  \item{x}{the text of the caption.
}
  \item{index}{integer from 1 to 9, defining the number of the heading style.
}
  \item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.}
}
\author{
Andri Signorell <andri@signorell.net>
}

\seealso{
\code{\link{ToWrd}}, \code{\link{WrdPlot}}, \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
}
\examples{
\dontrun{ # Windows-specific example
wrd <- GetNewWrd()

# insert a title in level 1
WrdCaption("My First Caption level 1", index=1, wrd=wrd)

# works as well for several levels
sapply(1:5, function(i)
  WrdCaption(gettextf("My First Caption level \%s",i), index=i, wrd=wrd)
)
}
}
\keyword{ print }
\concept{MS-Office}
