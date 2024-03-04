\name{WrdBookmark}
\alias{WrdInsertBookmark}
\alias{WrdDeleteBookmark}
\alias{WrdGoto}
\alias{WrdBookmark}
\alias{WrdUpdateBookmark}

\title{Some Functions to Handle MS-Word Bookmarks
}
\description{Accessing bookmarks by name is only possible by browsing the bookmark names. \code{WrdBookmark} returns a handle to a bookmark by taking its name as argument. \code{WrdInsertBookmark}, \code{WrdDeleteBookmark} inserts/deletes a bookmark in a Word document. \code{WrdGotoBookmark} allows to place the cursor on the bookmark and \code{WrdUpdateBookmark} replaces the content within the range of the bookmark in a Word document with the given text.
}
\usage{
WrdBookmark(name, wrd = DescToolsOfficeOptions("lastWord"))

WrdInsertBookmark(name, wrd = DescToolsOfficeOptions("lastWord"))
WrdDeleteBookmark(name, wrd = DescToolsOfficeOptions("lastWord"))

WrdGoto(name, what = wdConst$wdGoToBookmark, wrd = DescToolsOfficeOptions("lastWord"))

WrdUpdateBookmark(name, text, what = wdConst$wdGoToBookmark,
                  wrd = DescToolsOfficeOptions("lastWord"))
}
\arguments{
  \item{name}{the name of the bookmark.
}
  \item{text}{the text of the bookmark.
}
  \item{what}{a word constant, defining the type of object to be used to place the cursor.
}
  \item{wrd}{the pointer to a word instance. Can be a new one, created by \code{GetNewWrd()}
  or an existing one, created by \code{GetCurrWrd()}.
  Default is the last created pointer stored in \code{DescToolsOfficeOptions("lastWord")}.}
}

\details{Bookmarks are useful to build structured documents, which can be updated later.
}
\author{Andri Signorell <andri@signorell.net>
}


\seealso{
\code{\link{WrdFont}}, \code{\link{WrdPlot}}, \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
}
\examples{
\dontrun{ # we can't get this through the CRAN test - run it with copy/paste to console
wrd <- GetNewWrd()
WrdText("a)\n\n\nb)", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)
WrdInsertBookmark("chap_b")
WrdText("\n\n\nc)\n\n\n", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)

WrdGoto("chap_b")
WrdUpdateBookmark("chap_b", "Goto chapter B and set text")

WrdInsertBookmark("mybookmark")
ToWrd("A longer text\n\n\n")

# Now returning the bookmark
bm <- WrdBookmark("mybookmark")

# get the automatically created name of the bookmark
bm$name()
}}
\keyword{ print }
\concept{MS-Office}
