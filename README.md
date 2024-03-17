
<!-- README.md is generated from README.Rmd. Please edit that file -->
<!-- badges: start -->

[![CRAN
status](https://www.r-pkg.org/badges/version-last-release/DescToolsOffice)](https://CRAN.R-project.org/package=DescToolsOffice)
[![downloads](https://cranlogs.r-pkg.org/badges/grand-total/DescToolsOffice)](https://CRAN.R-project.org/package=DescToolsOffice)
[![downloads](http://cranlogs.r-pkg.org/badges/last-week/DescToolsOffice)](https://CRAN.R-project.org/package=DescToolsOffice)
[![License: GPL
v2+](https://img.shields.io/badge/License-GPL%20v2+-blue.svg)](https://www.gnu.org/licenses/old-licenses/gpl-2.0.en.html)
[![Lifecycle:
maturing](https://img.shields.io/badge/lifecycle-maturing-blue.svg)](https://lifecycle.r-lib.org/articles/stages.html)
[![R build
status](https://github.com/AndriSignorell/DescToolsOffice/workflows/R-CMD-check/badge.svg)](https://github.com/AndriSignorell/DescToolsOffice/actions)
[![pkgdown](https://github.com/AndriSignorell/DescToolsOffice/workflows/pkgdown/badge.svg)](https://andrisignorell.github.io/DescToolsOffice/)

<!-- badges: end -->

# Office Interface for DescTools

**DescToolsOffices** contains functions to produce documents using MS
Word (or PowerPoint) and functions to import data from Excel, based on
the functions contained in the package DescTools
(<https://CRAN.R-project.org/package=DescTools>).

Feedback, feature requests, bug reports and other suggestions are
welcome! Please report problems to [GitHub issues
tracker](https://github.com/AndriSignorell/DescTools/issues).

## Installation

You can install the released version of **DescToolsOffice** from
[CRAN](https://CRAN.R-project.org) with:

``` r
install.packages("DescToolsOffice")
```

And the development version from GitHub with:

``` r
if (!require("remotes")) install.packages("remotes")
remotes::install_github("AndriSignorell/DescToolsOffice")
```

# MS-Office

To make use of MS-Office features, you must have Office in one of its
variants installed. All `Wrd*`, `XL*` and `Pp*` functions require the
package **RDCOMClient** to be installed as well. Hence the use of these
functions is restricted to *Windows* systems. **RDCOMClient** can be
installed with:

``` r
install.packages("RDCOMClient", repos="http://www.omegahat.net/R")
```

The *omegahat* repository does not benefit from the same update service
as CRAN. So you may be forced to install a package compiled with an
earlier version, which usually is not a problem. Use e.g. for R 4.3.x:

``` r
url <- "http://www.omegahat.net/R/bin/windows/contrib/4.2/RDCOMClient_0.96-1.zip"
install.packages(url, repos = NULL, type = "binary")
```

**RDCOMClient** does not exist for Mac or Linux, sorry.

# Warning

**Warning:** This package is still under development. Although the code
seems meanwhile quite stable, until release of version 1.0 you should be
aware that everything in the package might be subject to change.
Backward compatibility is not yet guaranteed. Functions may be deleted
or renamed and new syntax may be inconsistent with earlier versions. By
release of version 1.0 the “deprecated-defunct process” will be
installed.

# Authors

Andri Signorell  
Helsana Versicherungen AG, Health Sciences, Zurich  
HWZ University of Applied Sciences in Business Administration Zurich.

R is a community project. This can be seen from the fact that this
package includes R source code and/or documentation previously published
by [various authors and
contributors](https://andrisignorell.github.io/DescToolsOffice/authors.html).
The good things come from all these guys, any problems are likely due to
my tweaking. Thank you all!

**Maintainer:** Andri Signorell

# Examples

``` r
library(DescToolsOffice)
```

<!-- ## Demo "describe" -->

``` r
demo(describe, package = "DescToolsOffice")
```
