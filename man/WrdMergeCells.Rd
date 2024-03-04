\name{WrdMergeCells}
\alias{WrdMergeCells}
\title{Merges Cells Of a Defined Word Table Range
}
\description{Merges a cell range of a word table.
}
\usage{
WrdMergeCells(wtab, rstart, rend)
}
\arguments{
  \item{wtab}{a handle to the word table as returned i.g. by  \code{\link{WrdTable}}
}
  \item{rstart}{the left/upper cell of the cell range.
}
  \item{rend}{the right/lower cell of the cell range.
}
}
\value{nothing
}
\author{Andri Signorell <andri@signorell.net>
}

\seealso{ \code{\link{WrdTable}}, \code{\link{WrdCellRange}}
}
\examples{
\dontrun{

# Windows-specific example
wrd <- GetNewWrd()
wtab <- WrdTable(nrow=3, ncol=3, wrd=wrd)
WrdMergeCells(wtab, rstart=c(1,2), rend=c(2,3))
}
}
\keyword{print}
\concept{MS-Office}
