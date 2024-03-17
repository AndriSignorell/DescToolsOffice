## Excel functions   ====





#' Use MS-Excel as Viewer for a Data.Frame
#' 
#' \code{XLView} can be used to view and edit a data.frame directly in
#' MS-Excel, resp. to create a new data.frame in MS-Excel.
#' 
#' The data.frame will be exported in CSV format and then imported in MS-Excel.
#' When importing data, MS-Excel will potentially change characters to numeric
#' values. If this seems undesirable (maybe we're loosing leading zeros) then
#' you should enclose the text in quotes and preset a =. x <-
#' \code{gettextf('="%s"', x)} would do the trick.  \cr\cr Take care: Changes
#' to the data made in MS-Excel will NOT automatically be updated in the
#' original data.frame. The user will have to read the csv-file into R again.
#' See examples how to get this done.\cr
#' 
#' \code{ToXL()} is used to export data frames or vectors directly to MS-Excel,
#' without export the data to a csv-file and import it on the XL side. So it it
#' possible to export several data.frames into one Workbook and edit the tables
#' after ones needs.
#' 
#' @aliases XLView 
#' @param x is a data.frame to be transferred to MS-Excel. If data is missing a
#' new file will be created.
#' @param row.names either a logical value indicating whether the row names of
#' x are to be written along with x, or a character vector of row names to be
#' written.
#' @param col.names either a logical value indicating whether the column names
#' of x are to be written along with x, or a character vector of column names
#' to be written.  See the section on 'CSV files' \code{\link{write.table}} for
#' the meaning of \code{col.names = NA}.
#' @param na the string to use for missing values in the data.
#' @param preserveStrings logical, will preserve strings from being converted
#' to numerics when imported in MS-Excel. See details. Default is \code{FALSE}.
#' @param sep the field separator string used for export of the object. Values
#' within each row of x are separated by this string.
#' @return the name/path of the temporary file edited in MS-Excel.
#' @note The function works only in Windows and requires \bold{RDCOMClient} to
#' be installed (see: Additional_repositories in DESCRIPTION of the package).
#' %% ~~further notes~~
#' @author Andri Signorell <andri@@signorell.net>, \code{ToXL()} is based on
#' code of Duncan Temple Lang <duncan@@r-project.org>
#' @seealso \code{\link{GetNewXL}}, \code{\link{XLGetRange}},
#' \code{\link{XLGetWorkbook}}
#' @keywords manip
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
#' @export XLView
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




#' Save Excel File
#' 
#' Save the current workbook under the given name and format.
#' 
#' 
#' @param fn the filename
#' @param file_format the file format using the xl constant.
#' @param xl the pointer to a MS-Excel instance. An new instance can be created
#' with \code{GetNewXL()}, returning the appropriate handle. A handle to an
#' already running instance is returned by \code{GetCurrXL()}.  Default is the
#' last created pointer stored in \code{DescToolsOptions("lastXL")}.
#' @return returns \code{TRUE} if the save operation has been successful
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{XLView}}
#' @keywords manip
#' @examples
#' 
#' \dontrun{# Windows-specific example
#' XLView(d.diamonds)
#' XLSaveAs("Diamonds")
#' xl$quit()
#' }
#' @export XLSaveAs
XLSaveAs <- function(fn, file_format=xlConst$XlFileFormat$xlWorkbookNormal, 
                     xl=DescToolsOptions("lastXL")){
  xl[["ActiveWorkbook"]]$SaveAs(FileName=fn, FileFormat=file_format)
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









#' Convert Excel Dates to POSIXct
#' 
#' As I repeatedly forgot how to convert Excel dates to POSIX here's the
#' specific function.
#' 
#' \code{\link{XLGetRange}} will return dates as integer values, because XL
#' stores them as integers. An Excel date can be converted with the (unusual)
#' origin of \code{as.Date(myDate, origin="1899-12-30")}, which is implemented
#' here.
#' 
#' Microsoft Excel supports two different date systems, the 1900 date system
#' and the 1904 date system. In the 1900 date system, the first day that is
#' supported is January 1, 1900. A date is converted into a serial number that
#' represents the number of elapsed days since January 1, 1900. In the 1904
#' date system, the first day that is supported is January 1, 1904. By default,
#' Microsoft Excel for the Macintosh uses the 1904 date system, Excel for
#' Windows the 1900 system. See also:
#' https://support.microsoft.com/en-us/kb/214330.
#' 
#' @param x the integer vector to be converted.
#' @param tz a time zone specification to be used for the conversion, if one is
#' required. See \code{\link{as.POSIXct}}.
#' @param xl1904 logical, defining if the unspeakable 1904-system should be
#' used. Default is FALSE.
#' @return return an object of the class POSIXct. Date-times known to be
#' invalid will be returned as NA.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{as.POSIXct}}
#' @keywords chron
#' @examples
#' 
#' XLDateToPOSIXct(41025)
#' XLDateToPOSIXct(c(41025.23, 41035.52))
#' 
#' @export XLDateToPOSIXct
XLDateToPOSIXct <- function (x, tz = "GMT", xl1904 = FALSE) {
  # https://support.microsoft.com/en-us/kb/214330
  if(xl1904)
    origin <- "1904-01-01"
  else
    origin <- "1899-12-30"
  
  as.POSIXct(x * (60 * 60 * 24), origin = origin, tz = tz)
}


###


