\name{GetNewXL}
\alias{GetNewXL}
\title{Create a New Excel Instance
}
\description{Start a new instance of Excel and return its handle.
This is needed to address the Excel application and objects afterwards.
}
\usage{
GetNewXL(visible = TRUE, newdoc = TRUE)
}
\arguments{
  \item{visible}{logical, should Excel made visible? Defaults to \code{TRUE}.
}
  \item{newdoc}{logical, determining if a new workbook should be created. Defaults to \code{TRUE}.
  }
}

\details{
Here's a list of some frequently used commands.\cr
Let's assume:
\preformatted{xl <- GetNewXL()
}
\tabular{ll}{
workbooks \tab \code{xl$workbooks()$count()} \cr
quit excel \tab \code{xl$quit()} \cr
}

}
\author{
Andri Signorell <andri@signorell.net>
}


\seealso{\code{\link{XLView}}, \code{\link{XLGetRange}}, \code{\link{XLGetWorkbook}}
}
\examples{
\dontrun{ # Windows-specific example
# get a handle to a new excel instance
xl <- GetNewXL()
}
}
\keyword{ misc }
