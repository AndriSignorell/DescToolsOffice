\name{WrdCellRange}
\alias{WrdCellRange}
\title{Return the Cell Range Of a Word Table
}
\description{Return a handle of a cell range of a word table. This is useful for
formating the cell range.
}
\usage{
WrdCellRange(wtab, from, to)
}
\arguments{
  \item{wtab}{a handle to the word table as returned i.g. by  \code{\link{WrdTable}}
}
  \item{from}{a vector containing row- and column number of the left/upper cell of the cell range.
}
  \item{to}{a vector containing row- and column number of the right/lower cell of the cell range.
}
}
\details{Cell range selecting might be complicated. This function makes it easy.
}
\value{a handle to the range.
}
\author{Andri Signorell <andri@signorell.net>
}

\seealso{ \code{\link{WrdTable}}
}
\examples{
\dontrun{

# Windows-specific example
wrd <- GetNewWrd()
WrdTable(nrow=3, ncol=3, wrd=wrd)
crng <- WrdCellRange(from=c(1,2), to=c(2,3))
crng$Select()
}
}
\keyword{print }
\concept{MS-Office}
