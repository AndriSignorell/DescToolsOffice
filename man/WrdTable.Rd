\name{WrdTable}
\alias{WrdTable}
\title{Insert a Table in a Word Document
}
\description{Create a table with a specified number of rows and columns in a Word document at the current position of the cursor.
}
\usage{
WrdTable(nrow = 1, ncol = 1, heights = NULL, widths = NULL, main = NULL,
         wrd = DescToolsOfficeOptions("lastWord"))
}
\arguments{
  \item{nrow}{number of rows.
}
  \item{ncol}{number of columns.
}
  \item{heights}{a vector of the row heights (in [cm]). If set to \code{NULL} (which is the default) the Word defaults will be used. The values will be recyled, if necessary.
}
  \item{widths}{a vector of the column widths (in [cm]). If set to \code{NULL} (which is the default) the Word defaults will be used. The values will be recyled, if necessary.
}
\item{main}{a caption for the plot. This will be inserted by InserCaption in Word. Default is NULL, which will insert nothing.}

\item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.
}
}
\value{A pointer to the inserted table.
}
\author{Andri Signorell <andri@signorell.net>
}


\seealso{\code{\link{GetNewWrd}}, \code{\link{ToWrd}}
}
\examples{
\dontrun{ # Windows-specific example
wrd <- GetNewWrd()
WrdTable(nrow=3, ncol=3, wrd=wrd)
}
}
\keyword{print }
\concept{MS-Office}
