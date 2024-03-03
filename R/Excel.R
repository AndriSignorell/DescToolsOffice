


## Excel functions   ====



XLView <- function (x, col.names = TRUE, row.names = FALSE, na = "", preserveStrings=FALSE, sep=";") {
  
  # # define some XL constants
  # xlToRight <- -4161
  
  fn <- paste(tempfile(pattern = "file", tmpdir = tempdir()),
              ".csv", sep = "")
  xl <- GetNewXL(newdoc=FALSE)
  owb <- xl[["Workbooks"]]
  
  if(!missing(x)){
    
    if(inherits(x, what = "ftable")){
      x <- FixToTable(capture.output(x), sep = " ", header = FALSE)
      col.names <- FALSE
    }
    
    if(preserveStrings){
      # embed all characters or factors in ="xyz"
      for(z in which(sapply(x, function(y) is.character(y) | is.factor(y)))){
        x[, z] <- gettextf('="%s', x[,z])
      }
    }
    
    write.table(x, file = fn, sep = sep, col.names = col.names,
                qmethod = "double", row.names = row.names, na=na)
    ob <- owb$Open(fn)
    # if row.names are saved there's the first cell in the first line missing
    # I don't actually see, how to correct this besides inserting a cell in XL
    if(row.names) xl$Cells(1, 1)$Insert(Shift=xlConst$xlToRight)
    xl[["Cells"]][["EntireColumn"]]$AutoFit()
    
  } else {
    owb$Add()
    awb <- xl[["ActiveWorkbook"]]
    # delete sheets(2,3) without asking, if it's ok
    xl[["DisplayAlerts"]] <- FALSE
    xl$Sheets(c(2,3))$Delete()
    xl[["DisplayAlerts"]] <- TRUE
    awb$SaveAs( Filename=fn, FileFormat=6 )
  }
  invisible(fn)
}


XLSaveAs <- function(fn, file_format=xlConst$XlFileFormat$xlWorkbookNormal, xl=DescToolsOfficeOptions("lastXL")){
  xl[["ActiveWorkbook"]]$SaveAs(FileName=fn, FileFormat=file_format)
}



ToXL <- function (x, at, ..., xl=DescToolsOfficeOptions("lastXL")) {
  stopifnot(IsValidHwnd(xl))   # "xl is not a valid Excel handle, use GetNewXL() or GetCurrXL().")
  UseMethod("ToXL")
}



ToXL.data.frame <- function(x, at, ..., xl=DescToolsOfficeOptions("lastXL"))
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



ToXL.matrix <- function (x, at, ..., xl = DescToolsOfficeOptions("lastXL")) {
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


ToXL.array <- function (x, at, ..., xl = DescToolsOfficeOptions("lastXL")) {
  
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



ToXL.table <- function (x, at, ..., xl = DescToolsOfficeOptions("lastXL")) {
  ToXL.array(x, at=at, ..., xl=xl)
}


ToXL.default <- function(x, at, byrow = FALSE, ..., xl=DescToolsOfficeOptions("lastXL")) {
  
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




XLCurrReg <- function(cell){
  structure(cell, class="XLCurrReg")
}


XLNamedReg <- function (x) {
  structure(x, class = "XLNamedReg")
}



XLColNames <- function() {
  c(LETTERS, out2 <- c(t(outer(LETTERS, LETTERS, paste, sep = ""))), 
    t(outer(LETTERS, out2, paste, sep = "")))[1:16384]
}



A1ToZ1S1 <- function(x){
  
  # was so slooow, we don't have to sort, if we do it a little more cleverly...
  # xlcol <- c( LETTERS
  #             , sort(c(outer(LETTERS, LETTERS, paste, sep="" )))
  #             , sort(c(outer(LETTERS, c(outer(LETTERS, LETTERS, paste, sep="" )), paste, sep="")))
  # )[1:16384]
  
  z1s1 <- function(x) {
    # remove all potential $ from a range first
    x <- gsub("\\$", "", x)
    colnr <- match( regmatches(x, regexec("^[[:alpha:]]+", x)), XLColNames())
    rownr <- as.numeric(regmatches(x, regexec("[[:digit:]]+$", x)))
    return(c(rownr, colnr))
  }
  
  lapply(unlist(strsplit(toupper(x),":")), z1s1)
}




XLGetRange <- function (file = NULL, sheet = NULL, range = NULL, as.data.frame = TRUE,
                        header = FALSE, stringsAsFactors = FALSE, echo = FALSE, 
                        na.strings = NULL, skip = 0) {
  
  # main function  *******************************
  
  # to do: 30.8.2015
  # we could / should check for a running XL instance here...
  # ans <- RDCOMClient::getCOMInstance("Excel.Application", force = FALSE, silent = TRUE)
  # if (is.null(ans) || is.character(ans)) print("not there")
  
  
  # https://stackoverflow.com/questions/38950005/how-to-manipulate-null-elements-in-a-nested-list/
  simple_rapply <- function(x, fn) {
    if(is.list(x)) {
      lapply(x, simple_rapply, fn)
    } else {
      fn(x)
    }
  }
  
  if(is.null(file)){
    xl <- GetCurrXL()
    ws <- xl$ActiveSheet()
    if(is.null(range)) {
      # if there is a selection in XL then use it, if only one cell selected use currentregion
      sel <- xl$Selection()
      if(sel$Cells()$Count() == 1 ){
        range <- xl$ActiveCell()$CurrentRegion()$Address(FALSE, FALSE)
      } else {
        range <- sapply(1:sel$Areas()$Count(), function(i) sel$Areas()[[i]]$Address(FALSE, FALSE) )
        
        # old: this did not work on some XL versions with more than 28 selected areas
        # range <- xl$Selection()$Address(FALSE, FALSE)
        # range <- unlist(strsplit(range, ";"))
        # there might be more than 1 single region, split by ;
        # (this might be a problem for other locales)
      }
    }
    
  } else {
    xl <- GetNewXL()
    wb <- xl[["Workbooks"]]$Open(file)
    
    # set defaults for sheet and range here
    if(is.null(sheet))
      sheet <- 1
    
    if(is.null(range))
      range <- xl$Cells(1,1)$CurrentRegion()$Address(FALSE, FALSE)
    
    ws <- wb$Sheets(sheet)$select()
  }
  
  if(inherits(x=range, what="XLCurrReg")){
    # take only the first cell of a given range
    zs <- A1ToZ1S1(range)[[1]]
    range <- xl$Cells(zs[1], zs[2])$CurrentRegion()$Address(FALSE, FALSE)
  } else if(inherits(x=range, what="XLNamedReg")){
    # get the address of the named region
    sel <- xl$ActiveWorkbook()$Names(as.character(range))$RefersToRange()
    range <- sapply(1:sel$Areas()$Count(), function(i) sel$Areas()[[i]]$Address(FALSE, FALSE) )
    
  }
  
  # recycle skip
  skip <- rep(skip, length.out=length(range))
  
  lst <- list()
  for (i in seq_along(range)) {
    zs <- A1ToZ1S1(range[i])
    if(length(zs)==1){
      rr <- xl$Cells(zs[[1]][1], zs[[1]][2])
    } else {
      rr <- xl$Range(xl$Cells(zs[[1]][1], zs[[1]][2]), xl$Cells(zs[[2]][1], 
                                                                zs[[2]][2]))
    }
    
    # resize and offset range, if skip != 0
    if (skip[i] != 0) 
      rr <- rr$Resize(rr$Rows()$Count() - skip[i])$Offset(skip[i], 0)
    
    # Get the values
    if(is.null(rr[["Value"]]))
      # this is the case when we have multiple ranges selected an one of them 
      # is a single empty cell
      lst[[i]] <- NA
    else 
      lst[[i]] <- rr[["Value"]]
    # this produces a non trappable warning "Unhandled conversion type 10"
    # no further problem, but document in help!
    
    if(!is.list(lst[[i]]))
      lst[[i]] <- list(as.list(lst[[i]]))
    
    # replace NULLs by NAs (rather complicated job...)
    lst[[i]] <- simple_rapply(lst[[i]], 
                              function(x) if(is.null(x)) NA else x)
    
    # # address of errors: rr$SpecialCells(xlConst$xlFormulas, xlConst$xlErrors)$address()
    lst[[i]] <- rapply(lst[[i]],
                       function(x) {
                         
                         if(inherits(x=x, what="VARIANT")){
                           # if there are errors replace them by NA
                           NA
                           
                         } else if(inherits(x=x, what="COMDate")) {
                           # if there are XL dates, replace them by their date value
                           if(IsWhole(x))
                             as.Date(XLDateToPOSIXct(x))
                           else
                             XLDateToPOSIXct(x)
                           
                         } else if(x %in% na.strings) {
                           # if x in na.strings' list replace it by NA
                           NA
                           
                         } else {  
                           x
                         }
                       }, how = "replace")
    
    names(lst)[i] <- range[i]
  }
  
  if (as.data.frame) {
    for (i in seq_along(lst)) {
      
      if (header) {
        xnames <- unlist(lapply(lst[[i]], "[", 1))
        lst[[i]] <- lapply(lst[[i]], "[", -1)
      }
      
      # This was old: not fall back to it!!
      # lst[[i]] <- do.call(data.frame, c(lapply(lst[[i]][], 
      #                                          unlist), stringsAsFactors = stringsAsFactors))
      
      # don't use lapply and unlist as it's killing the classes for dates
      # https://stackoverflow.com/questions/15659783/why-does-unlist-kill-dates-in-r
      lst[[i]] <- do.call(data.frame, c(
        lapply(lst[[i]], function(x) do.call(c, x)), 
        stringsAsFactors = stringsAsFactors))
      
      if (header) {
        names(lst[[i]]) <- xnames
        
      } else {
        names(lst[[i]]) <- paste("X", 1:ncol(lst[[i]]), sep = "")
      }
    }
  }
  
  # just return a single object (for instance data.frame) if only one range was supplied
  if (length(lst) == 1)   lst <- lst[[1]]
  
  attr(lst, "call") <- gettextf("XLGetRange(file = %s, sheet = %s,\n     range = c(%s),\n     as.data.frame = %s, header = %s, stringsAsFactors = %s)", 
                                gsub("\\\\", "\\\\\\\\", shQuote(paste(xl$ActiveWorkbook()$Path(), 
                                                                       xl$ActiveWorkbook()$Name(), sep = "\\"))), shQuote(xl$ActiveSheet()$Name()), 
                                gettextf(paste(shQuote(range), collapse = ",")), as.data.frame, 
                                header, stringsAsFactors)
  
  if (!is.null(file)) {
    xl$ActiveWorkbook()$Close(savechanges = FALSE)
    xl$Quit()                  # only quit, if a new XL-instance was created before
  }
  
  if (echo) 
    cat(attr(lst, "call"))
  
  class(lst) <- c("xlrange", class(lst))
  return(lst)
  
}



as.matrix.xlrange <- function(x, ...){
  SetNames(as.matrix(x[[1]]), rownames=x[[2]][,1], colnames=x[[3]][1,])
}



XLGetWorkbook <- function (file, compactareas = TRUE) {
  
  
  IsEmptySheet <- function(sheet)
    sheet$UsedRange()$Rows()$Count() == 1 &
    sheet$UsedRange()$columns()$Count() == 1 &
    is.null(sheet$cells(1,1)$Value())
  
  CompactArea <- function(lst)
    do.call(cbind, lapply(lst, cbind))
  
  
  # xlCellTypeConstants <- 2
  # xlCellTypeFormulas <- -4123
  
  xl <- GetNewXL()
  wb <- xl[["Workbooks"]]$Open(file)
  
  lst <- list()
  for (i in 1:wb$Sheets()$Count()) {
    
    if(!IsEmptySheet(sheet=xl$Sheets(i))) {
      
      # has.formula is TRUE, when all cells contain formula, FALSE when no cell contains a formula
      # and NULL else, thus: !identical(FALSE) for having some or all
      if(!identical(xl$Sheets(i)$UsedRange()$HasFormula(), FALSE))
        areas <- xl$union(
          xl$Sheets(i)$UsedRange()$SpecialCells(xlConst$xlCellTypeConstants),
          xl$Sheets(i)$UsedRange()$SpecialCells(xlConst$xlCellTypeFormulas))$areas()
      else
        areas <- xl$Sheets(i)$UsedRange()$SpecialCells(xlConst$xlCellTypeConstants)$areas()
      
      alst <- list()
      for ( j in 1:areas$count())
        alst[[j]] <- areas[[j]]$Value2()
      
      lst[[xl$Sheets(i)$name()]] <- alst
      
    }
  }
  
  if(compactareas)
    lst <- lapply(lst, function(x) lapply(x, CompactArea))
  
  # close without saving
  wb$Close(FALSE)
  
  xl$Quit()
  return(lst)
  
}



XLKill <- function(){
  # Excel would only quit, when all workbooks are closed before, someone said.
  # http://stackoverflow.com/questions/15697282/excel-application-not-quitting-after-calling-quit
  
  # We experience, that it would not even then quit, when there's no workbook loaded at all.
  # maybe gc() would help ??
  # so killing the task is "ultima ratio"...
  
  shell('taskkill /F /IM EXCEL.EXE')
}



XLDateToPOSIXct <- function (x, tz = "GMT", xl1904 = FALSE) {
  # https://support.microsoft.com/en-us/kb/214330
  if(xl1904)
    origin <- "1904-01-01"
  else
    origin <- "1899-12-30"
  
  as.POSIXct(x * (60 * 60 * 24), origin = origin, tz = tz)
}


###


