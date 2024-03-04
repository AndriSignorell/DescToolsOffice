
# XL GetRange



#' Import Data Directly From Excel
#' 
#' The package \code{RDCOMClient} is used to open an Excel workbook and return
#' the content (value) of one (or several) given range(s) in a specified sheet.
#' This is helpful, whenever pathologically scattered data on an Excel sheet,
#' which can't simply be saved as CSV-file, has to be imported in R.\cr\cr
#' \code{XLGetWorkbook()} does the same for all the sheets in an Excel
#' workbook.
#' 
#' The result consists of a list of lists, if \code{as.data.frame} is set to
#' \code{FALSE}. Be then prepared to encounter \code{NULL} values. Those will
#' prevent from easily being able to coerce the square data structure to a
#' data.frame.
#' 
#' The following code will replace the \code{NULL} values by \code{NA} and
#' coerce the data to a data.frame. \preformatted{ # get the range D1:J69 from
#' an excel file xlrng <- XLGetRange(file="myfile.xlsx", sheet="Tabelle1",
#' range="D1:J69", as.data.frame=FALSE)
#' 
#' # replace NULL values by NA xlrng[unlist(lapply(xlrng, is.null))] <- NA
#' 
#' # coerce the square data structure to a data.frame d.lka <-
#' data.frame(lapply(data.frame(xlrng), unlist)) } This of course can be
#' avoided by setting \code{as.data.frame} = \code{TRUE}.
#' 
#' The function will return dates as integers, because MS-Excel stores them
#' internally as integers. Such a date can subsequently be converted with the
#' (unusual) origin of \code{as.Date(myDate, origin="1899-12-30")}. See also
#' \code{\link{XLDateToPOSIXct}}, which does the job. The conversion can
#' directly be performed by \code{XLGetRange()} if \code{datecols} is used and
#' contains the date columns in the sheet data.
#' 
#' @aliases XLGetRange XLGetWorkbook XLCurrReg XLNamedReg

#' @param file the fully specified path and filename of the workbook. If it is
#' left as \code{NULL}, the function will look for a running Excel-Application
#' and use its current sheet. The parameter \code{sheet} will be ignored in
#' this case.
#' @param sheet the name of the sheet containing the range(s) of interest.
#' @param range a scalar or a vector with the address(es) of the range(s) to be
#' returned (characters).  Use "A1"-address mode to specify the ranges, for
#' example \code{"A1:F10"}. \cr If set to \code{NULL} (which is the default),
#' the function will look for a selection that contains more than one cell. If
#' found, the function will use this selection. If there is no selection then
#' the current region of the selected cell will be used. Use \code{XLCurrReg()}
#' if the current region of a cell, which is currently not selected, should be
#' used. Range names can be provided with \code{XLNamedReg("name")}.
#' @param as.data.frame logical. Determines if the cellranges should be coerced
#' into data.frames. Defaults to \code{TRUE}, as this is probably the common
#' use of this function.
#' @param header a logical value indicating whether the range contains the
#' names of the variables as its first line. Default is \code{FALSE}.
#' \code{header} is ignored if \code{as.data.frame} has been set to
#' \code{FALSE}.
#' @param stringsAsFactors logical. Should character columns be coerced to
#' factors? The default is \code{FALSE}, which will return character vectors.
#' @param echo logical. If set to \code{TRUE}, the function will print the full
#' command used, such that it can be copied into the R-script for future use.
#' @param na.strings a character vector of strings which are to be interpreted
#' as \code{NA} values. Blank fields are always considered to be missing
#' values. Default is \code{NULL}, meaning none.
#' @param compactareas logical, defining if areas should be returned by
#' \code{XLGetWorkbook} as list or as matrix (latter is default).
#' @param cell range of the left uppe cell, when current region should be used.
#' @param x the name or the index of the XL-name to be used.
#' @param skip the number of lines of the data file to skip before beginning to
#' read data.

#' @return If \code{as.data.frame} is set to \code{TRUE}, a single data.frame
#' or a list of data.frames will be returned. If set to \code{FALSE} a list of
#' the cell values in the specified Excel range, resp. a list of lists will be
#' returned.
#' 
#' \code{XLGetWorkbook()} returns a list of lists of the values in the given
#' workbook.

#' @author Andri Signorell <andri@@signorell.net>

#' @seealso \code{\link{GetNewXL}}, \code{\link{GetCurrXL}},
#' \code{\link{XLView}}

#' @keywords manip

#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' 
#' XLGetRange(file="C:/My Documents/data.xls",
#'            sheet="Sheet1",
#'            range=c("A2:B5","M6:X23","C4:D40"))
#' 
#' 
#' # if the current region has to be read (incl. a header), place the cursor in the interesting region
#' # and run:
#' d.set <- XLGetRange(header=TRUE)
#' 
#' # Get XL nameslist
#' nm <- xl$ActiveWorkbook()$names()
#' 
#' lst <- list()
#' for(i in 1:nm$count())
#'   lst[[i]] <- c(name=nm[[i]]$name(), 
#'                 address=nm[[i]]$refersToRange()$Address())
#'   
#' # the defined names
#' as.data.frame(do.call(rbind, lst), stringsAsFactors = FALSE)
#' }
#' 
#' @export XLGetRange
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


#' @rdname XLGetRange
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



#' @rdname XLGetRange
XLCurrReg <- function(cell){
  structure(cell, class="XLCurrReg")
}


#' @rdname XLGetRange
XLNamedReg <- function (x) {
  structure(x, class = "XLNamedReg")
}


