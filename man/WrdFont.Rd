\name{WrdFont}
\alias{WrdFont}
\alias{WrdFont<-}
\title{Get or Set the Font in Word
}
\description{\code{WrdFont} can be used to get and set the font in Word for the text to be inserted. \code{WrdFont} returns the font at the current cursor position.
}
\usage{
WrdFont(wrd = DescToolsOfficeOptions("lastWord"))
WrdFont(wrd) <- value
}
\arguments{
  \item{value}{the font to be used to the output. This should be defined as a list containing fontname, fontsize, bold and italic flags:\cr \code{list(name="Arial", size=10, bold=FALSE, italic=TRUE, color=wdConst$wdColorBlack)}.
}
  \item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.}

}
\details{The font color can be defined by a Word constant beginning with \code{wdConst$wdColor}.
The defined colors can be listed with \code{grep("wdColor", names(wdConst), val=TRUE)}.
}
\value{
a list of the attributes of the font in the current cursor position:
  \item{name}{the fontname}
  \item{size}{the fontsize}
  \item{bold}{bold}
  \item{italic}{italic}
  \item{color}{the fontcolor}

}
\author{
Andri Signorell <andri@signorell.net>
}

\seealso{\code{\link{ToWrd}}, \code{\link{WrdPlot}}, \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
}
\examples{
\dontrun{ # Windows-specific example

wrd <- GetNewWrd()

for(i in seq(10, 24, 2))
  ToWrd(gettextf("This is Arial size \%s \n", i), font=list(name="Arial", size=i))

for(i in seq(10, 24, 2))
  ToWrd(gettextf("This is Times size \%s \n", i), font=list(name="Times", size=i))
}}
\keyword{ print }
\concept{MS-Office}
