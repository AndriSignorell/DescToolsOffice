\name{GetCurrWrd}
\alias{GetCurrWrd}
\alias{GetCurrXL}

\title{Get a Handle to a Running Word/Excel Instance

}
\description{Look for a running Word, resp. Excel instance and return its handle. If no running instance is found a new instance will be created (which will be communicated with a warning).

}
\usage{
GetCurrWrd()
GetCurrXL()
}


\value{a handle (pointer) to the running Word, resp. Excel instance. 
}
\note{When closing an application instance, the value of the pointer in R is not somehow automatically invalidated. In such cases the corresponding variable contains an invalid address. 
Whether the pointer still refers to a valid running application instance can be checked by \code{\link{IsValidHwnd}}. }


\author{
Andri Signorell <andri@signorell.net>
}

\seealso{\code{\link{GetNewWrd}}, \code{\link{IsValidHwnd}}
}
\examples{
\dontrun{# Windows-specific example

# Start a new instance
GetNewWrd()

# grab the handle to this instance
wrd <- GetCurrWrd()

# this should be valid
IsValidHwnd(wrd)

# close the instance
wrd$quit()

# now it should be gone and the pointer invalid
if(IsValidHwnd(wrd)){ 
  print("Ouups! Still there?")
} else {  
  print("GetCurrWrd: no running word instance found...")
}
}
}
\keyword{ misc }
