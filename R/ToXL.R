
# ToXL

#' Transfer Data to Excel
#' 
#' \code{ToXL()} is used to export data frames or vectors directly to MS-Excel,
#' without export the data to a csv-file and import it on the XL side. So it it
#' possible to export several data.frames into one Workbook and edit the tables
#' after ones needs.
#' 
#' @aliases ToXL 
#' @param x is a data.frame to be transferred to MS-Excel. If data is missing a
#' new file will be created.
#' @param at can be a range adress as character (e.g. \code{"A1"}), a vector of
#' 2 integers (e.g \code{c(1,1)}) or a cell object as it is returned by
#' \code{xl$Cells(1,1)}, denominating the left upper cell, where the data.frame
#' will be placed in the MS-Excel sheet.
#' @param xl the pointer to a MS-Excel instance. An new instance can be created
#' with \code{GetNewXL()}, returning the appropriate handle. A handle to an
#' already running instance is returned by \code{GetCurrXL()}.  Default is the
#' last created pointer stored in \code{DescToolsOptions("lastXL")}.
#' @param \dots further arguments are not used.

#' @return the name/path of the temporary file edited in MS-Excel.

#' @note The function works only in Windows and requires \bold{RDCOMClient} to
#' be installed (see: Additional_repositories in DESCRIPTION of the package).
#' 
#' @author Andri Signorell <andri@@signorell.net>, \code{ToXL()} is based on
#' code of Duncan Temple Lang <duncan@@r-project.org>
#' 
#' @seealso \code{\link{GetNewXL}}, \code{\link{XLGetRange}},
#' \code{\link{XLGetWorkbook}}
#' 
#' @keywords manip
#' 
#' @examples
#' 
#' \dontrun{
#' # Windows-specific example
#' XLView(d.diamonds)
#' 
#' # edit an existing data.frame in MS-Excel, make changes and save there, return the filename
#' fn <- XLView(d.diamonds)
#' # read the changed file and store in new data.frame
#' d.frm <- read.table(fn, header=TRUE, quote="", sep=";")
#' 
#' # Create a new file, edit it in MS-Excel...
#' fn <- XLView()
#' # ... and read it into a data.frame when in R again
#' d.set <- read.table(fn, header=TRUE, quote="", sep=";")
#' 
#' # Export a ftable object, quite elegant...
#' XLView(format(ftable(Titanic), quote=FALSE), row.names = FALSE, col.names = FALSE)
#' 
#' 
#' # Export a data.frame directly to XL, combined with subsequent formatting
#' 
#' xl <- GetNewXL()
#' owb <- xl[["Workbooks"]]$Add()
#' sheet <- xl$Sheets()$Add()
#' sheet[["name"]] <- "pizza"
#' 
#' ToXL(d.pizza[1:10, 1:10], xl$Cells(1,1))
#' 
#' obj <- xl$Cells()$CurrentRegion()
#' obj[["VerticalAlignment"]] <- xlConst$xlTop
#' 
#' row <- xl$Cells()$CurrentRegion()$rows(1)
#' # does not work:   row$font()[["bold"]] <- TRUE
#' # works:
#' obj <- row$font()
#' obj[["bold"]] <- TRUE
#' 
#' obj <- row$borders(xlConst$xlEdgeBottom)
#' obj[["linestyle"]] <- xlConst$xlContinuous
#' 
#' cols <- xl$Cells()$CurrentRegion()$columns(1)
#' cols[["HorizontalAlignment"]] <- xlConst$xlLeft
#' 
#' xl$Cells()$CurrentRegion()[["EntireColumn"]]$AutoFit()
#' cols <- xl$Cells()$CurrentRegion()$columns(4)
#' cols[["WrapText"]] <- TRUE
#' cols[["ColumnWidth"]] <- 80
#' xl$Cells()$CurrentRegion()[["EntireRow"]]$AutoFit()
#' 
#' sheet <- xl$Sheets()$Add()
#' sheet[["name"]] <- "whisky"
#' ToXL(d.whisky[1:10, 1:10], xl$Cells(1,1))}
#' 
#' @export ToXL

ToXL <- function (x, at, ..., xl=DescToolsOptions("lastXL")) {
  stopifnot(IsValidHwnd(xl))   # "xl is not a valid Excel handle, use GetNewXL() or GetCurrXL().")
  UseMethod("ToXL")
}



ToXL.data.frame <- function(x, at, ..., xl=DescToolsOptions("lastXL"))
  ## export the data.frame "x" into the location "at" (top,left cell)
  ## output the occupying range.
  ## TODO: row.names, more error checking
{
  if(is.character(at)){
    # address of the left upper cell
    at <- do.call(xl$Cells, as.list(A1ToZ1S1(at)[[1]]))
    
  } else if(is.vector(at)) {
    # get a handle of the cell range
    at <- do.call(xl$Cells, as.list(at))
  }
  
  nc <- dim(x)[2]
  if(nc < 1) stop("data.frame must have at least one column")
  r1 <- at$Row()                   ## 1st row in range
  c1 <- at$Column()                ## 1st col in range
  c2 <- c1 + nc - 1                ## last col (*not* num of col)
  ws <- at[["Worksheet"]]
  
  ## headers
  if(!is.null(names(x))) {
    hdrRng <- ws$Range(ws$Cells(r1, c1), ws$Cells(r1, c2))
    hdrRng[["Value"]] <- names(x)
    rng <- ws$Cells(r1 + 1, c1)
  } else {
    rng <- ws$Cells(r1, c1)
  }
  
  ## data
  for(j in seq(from = 1, to = nc)){
    # debug only:
    # cat("Column", j, "\n")
    ToXL(x[, j], at = rng, xl=xl)   ## no byrow for data.frames!
    rng <- rng$Next()               ## next cell to the right
  }
  invisible(ws$Range(ws$Cells(r1, c1), ws$Cells(r1 + nrow(x), c2)))
}



ToXL.matrix <- function (x, at, ..., xl = DescToolsOptions("lastXL")) {
  ## export the matrix "x" into the location "at" (top,left cell)
  
  if(is.character(at)){
    # address of the left upper cell
    at <- do.call(xl$Cells, as.list(A1ToZ1S1(at)[[1]]))
    
  } else if(is.vector(at)) {
    # get a handle of the cell range
    at <- do.call(xl$Cells, as.list(at))
  }
  
  nc <- dim(x)[2]
  if (nc < 1) 
    stop("matrix must have at least one column")
  
  if(!is.null(names(dimnames(x)))) {
    ToXL(names(dimnames(x))[1], at=at$offset(1, 0)$address())
    fnt <- at$offset(1, 0)$Font()
    fnt[["Bold"]] <- TRUE
    ToXL(dimnames(x)[[1]], at=at$offset(2, 0)$address())
    at_rn <- at$offset(2, 0)$resize(length(dimnames(x)[[1]]), 1)
    at_rn[["IndentLevel"]] <- 1
    ToXL(names(dimnames(x))[2], at=at$offset(0, 1)$address())
    fnt <- at$offset(0, 1)$Font()
    fnt[["Bold"]] <- TRUE
    ToXL(rbind(dimnames(x)[[2]]), at=at$offset(1, 1)$address())
    at <- at$offset(2, 1)
  }
  
  xref <- RDCOMClient::asCOMArray(x)
  rng <- at$resize(dim(x)[1], dim(x)[2])
  rng[["Value"]] <- xref
  
  invisible(rng)
  
}


ToXL.array <- function (x, at, ..., xl = DescToolsOptions("lastXL")) {
  
  if(is.character(at)){
    # address of the left upper cell
    at <- do.call(xl$Cells, as.list(A1ToZ1S1(at)[[1]]))
    
  } else if(is.vector(at)) {
    # get a handle of the cell range
    at <- do.call(xl$Cells, as.list(at))
  }
  
  lst <- lapply(asplit(x, seq_along(dim(x))[-c(1:2)]), "[")
  
  g <- expand.grid(dimnames(x)[-c(1:2)])
  names(lst) <- paste0(", , ", apply(sapply(colnames(g), function(x) paste(x, "=", g[, x])), 1, paste, collapse=", "))
  
  for(i in seq_along(lst)){
    ToXL(names(lst)[i], at=at)
    at <- at$offset(2, 0)
    ToXL(lst[[i]], at=at)
    at <- at$offset(dim(lst[[i]])[1] + 3, 0)
  }
  
  
}



ToXL.table <- function (x, at, ..., xl = DescToolsOptions("lastXL")) {
  ToXL.array(x, at=at, ..., xl=xl)
}


ToXL.default <- function(x, at, byrow = FALSE, ..., xl=DescToolsOptions("lastXL")) {
  
  #  function(x, at = NULL, byrow = FALSE, ...)
  ## coerce x to a simple (no attributes) vector and export to
  ## the range specified at "at" (can refer to a single starting cell);
  ## byrow = TRUE puts x in one row, otherwise in one column.
  ## How should we deal with unequal of ranges and vectors?  Currently
  ## we stop, modulo the special case when at refers to the starting cell.
  ## TODO: converters (currency, dates, etc.)
  
  if(is.character(at)){
    # address of the left upper cell
    at <- do.call(xl$Cells, as.list(A1ToZ1S1(at)[[1]]))
    
  } else if(is.vector(at)) {
    # get a handle of the cell range
    at <- do.call(xl$Cells, as.list(at))
  }
  
  n <- length(x)
  if(n < 1) return(at)
  d <- c(at$Rows()$Count(), at$Columns()$Count())
  N <- prod(d)
  
  xl <- at$Application()
  
  if(N == 1 && n > 1){     ## at refers to the starting cell
    r1c1 <- c(at$Row(), at$Column())
    r2c2 <- r1c1 + if(byrow) c(0, n-1) else c(n-1, 0)
    ws <- at[["Worksheet"]]
    at <- ws$Range(ws$Cells(r1c1[1], r1c1[2]),
                   ws$Cells(r2c2[1], r2c2[2]))
  } else if(n != N)
    stop("range and length(x) differ")
  
  ## currently we can only export primitives...
  
  if(any(class(x) %in% c("logical", "integer", "numeric", "character")))
    x <- as.vector(x)     ## clobber attributes
  
  else
    x <- as.character(x)  ## give up -- coerce to chars
  
  ## here we create a C-level COM safearray
  d <- if(byrow) c(1, n) else c(n, 1)
  # is this an alternative??
  # RDCOMClient::asCOMArray(matrix(x, nrow=d[1], ncol=d[2]))
  #  xref <- .Call("R_create2DArray", PACKAGE="RDCOMClient", matrix(x, nrow=d[1], ncol=d[2]))
  xref <- RDCOMClient::asCOMArray(matrix(x, nrow=d[1], ncol=d[2]))
  at[["Value"]] <- xref
  
  # workaround for missing values, simply delete the transferred bullshit
  na <- which(is.na(x))
  if(length(na) > 0) {
    if(byrow){
      arow <- gsub("[A-Z]","", at$cells(1,1)$address(rowabsolute=FALSE, columnabsolute=FALSE))
      
      # xlcol <- c( LETTERS
      #             , sort(c(outer(LETTERS, LETTERS, paste, sep="" )))
      #             , sort(c(outer(LETTERS, c(outer(LETTERS, LETTERS, paste, sep="" )), paste, sep="")))
      # )[1:16384]
      # xlcol <- XLColNames
      rngA1 <- paste(XLColNames()[na], arow, sep="", collapse = ";")
      rng <- xl$range(rngA1)$offset(ColumnOffset=xl$Range(at$Address())$Column()-1)
      
    } else {
      # find the column
      acol <- gsub("[0-9]","", at$cells(1,1)$address(rowabsolute=FALSE, columnabsolute=FALSE))
      # build range adress for the NAs
      rngA1 <- paste(acol, na, sep="", collapse = ";")
      # offset, if there's a name
      rng <- xl$range(rngA1)$offset(xl$Range(at$Address())$Row()-1)
    }
    rng[["FormulaR1C1"]] <- ""
  }
  
  invisible(at)
}

