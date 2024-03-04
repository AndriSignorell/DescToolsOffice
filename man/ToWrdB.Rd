\name{ToWrdB}
\alias{ToWrdB}
\title{Send Objects to Word and Bookmark Them
}
\description{Send objects like tables, ftables, lm tables, TOnes or just simple texts to a MS-Word document and place a bookmark on them. This has the advantage, that objects in a Word document can be updated later, provided the bookmark name has been stored.
}
\usage{
ToWrdB(x, font = NULL, ..., wrd = DescToolsOfficeOptions("lastWord"), 
        bookmark = gettextf("bmt\%s", round(runif(1, min = 0.1) * 1e+09)))
}
\arguments{
  \item{x}{the object to be transferred to Word.
}
  \item{font}{the font to be used to the output. This should be defined as a list containing fontname, fontsize, bold and italic flags:\cr \code{list(name="Arial", size=10, bold=FALSE, italic=TRUE)}.
}
   \item{\dots}{further arguments to be passed to or from methods.
}
 \item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.
}
  \item{bookmark}{the name of the bookmark.
}
}
\details{This function encapsulates \code{\link{ToWrd}}, by placing a bookmark over the complete inserted results.
The given name can be questioned with \code{bm$name()}.
}
\value{a handle to the set bookmark
}
\author{Andri Signorell <andri@signorell.net>
}


\seealso{\code{\link{ToWrd}}, \code{\link{WrdInsertBookmark}}
}
\examples{
\dontrun{
# we can't get this through the CRAN test - run it with copy/paste to console

wrd <- GetNewWrd()
bm <- ToWrdB("This is text to be possibly replaced later.")

# get the automatically created name of the bookmark
bm$name()

WrdGoto(bm$name())
UpdateBookmark(...)
}
}
\keyword{ print }

