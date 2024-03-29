% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/GetHandle.R
\name{WrdKill}
\alias{WrdKill}
\alias{XLKill}
\title{End Application Task}
\usage{
WrdKill()

XLKill()
}
\description{
\code{WrdKill()} and \code{XLKill()} will shut down a running application instance
(which also might be invisible).
Background is the fact, that the simple quit() command not always
terminates a running XL task, and only sets it invisible (which can be
observed the TaskManager).
This ghost instance may sometimes confuse XLView and hinder to
create a new instance. In such cases we have to do the garbage
collection and "killing" the process seems ultima ratio.
}
\examples{
\dontrun{ # Windows-specific example
# get a handle to a new Word instance
wrd <- GetNewWrd()
# end it with the crowbar 
WrdKill()

# get a handle to a new Excel instance
xl <- GetNewXL()
# end it with the crowbar 
XLKill()
}

}
\seealso{
\code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
}
\author{
Andri Signorell \href{mailto:andri@signorell.net}{andri@signorell.net}
}
\keyword{misc}
