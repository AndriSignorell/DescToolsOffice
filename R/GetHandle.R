

#' Get a Handle to COM Application
#' 
#' To be able to control COM applications we need a so called handle,
#' a kind of address, to it. This handle can be grabbed with the underlying functions
#' from the package \code{RDCOMClient}.
#' 
#' The function \code{GetCOMAppHandle()} is the workhorse for 
#' launching a new instance of the required application and return its handle.
#' It checks if the RDCOMClient package is available and tries to start the
#' application. If this is successful the handle is stored as option with the given name
#' and returned as result.
#'  
#' 
#' @aliases GetCOMAppHandle createCOMReference

#' @param ref	the S object that is an external pointer containing the reference to the COM object.
#' @param className	the name of the class that is “suggested” by the caller.

#' @param app name of the application as required by COM. This is typically the name of 
#' the program extended by \code{".Application"}, e.g. for MS-Word it would be 
#' \code{"Word.Application"}. 
#' @param option logical, should the handle be stored as an option? 
#' If this is left to \code{NULL} nothing will be stored, if a text is provided, the
#' handle will be stored under this name.
#' @param existing logical, should the handle to an already existing instance be returned? 
#' Defaults to \code{FALSE}.
#' @param visible logical, should the application made visible? Defaults to \code{TRUE}.

#' @return a handle (pointer) to a running Word/Excel/PowerPoint or another 
#' COM instance.

#' @author Andri Signorell <andri@@signorell.net>

#' @seealso \code{\link{GetHwnd}}, \code{\link{GetCurrWrd}}, \code{\link{GetCurrXL}}, 
#' \code{\link{GetCurrPP}}, \code{\link{IsValidHwnd}}
#' 
#' @keywords misc
#' @examples
#' 
#' \dontrun{# Windows-specific example
#' 
#' # Start a new instance of an application and store the handle as option "lastWord"
#' hwnd <- GetCOMAppHandle("Word.Application", option = "lastWord", 
#'                          existing = FALSE, visible = TRUE)
#' DescToolsOptions("lastWord")
#' 
#' # close the application
#' hwnd$quit()
#' }

#' @export createCOMReference
createCOMReference <- function(ref, className) {
  RDCOMClient::createCOMReference(ref, className)
}


#' @rdname createCOMReference
#' @export GetCOMAppHandle
GetCOMAppHandle <- function(app, option=NULL, existing=FALSE, visible=TRUE){
  
  if (requireNamespace("RDCOMClient", quietly = FALSE)) {
    
    if(!existing)
      # there's no "get"-function in RDCOMClient, so just create a new here..
      hwnd <- RDCOMClient::COMCreate(app, force=TRUE, existing=FALSE)
    else
      hwnd <- RDCOMClient::getCOMInstance(app, force=FALSE, silent=TRUE)
    
    if(!(!is.null(hwnd) && !is.character(hwnd))){
      warning(gettextf("No running %s application found!", app))
      return(hwnd)
    } 
    
    else
      if(visible) 
        hwnd[["Visible"]] <- visible
      
      
    # store the handle as option, if required
    if(!is.null(option))
        eval(parse(text=gettextf("DescToolsOptions(%s = hwnd)", option)))

  } else {
    
    # no RDCOMClient present or not Windows system
    if(Sys.info()["sysname"] == "Windows")
      warning("RDCOMClient is not available. To install it use: install.packages('RDCOMClient', repos = 'http://www.stats.ox.ac.uk/pub/RWin/')")
    else
      warning(gettextf("RDCOMClient is unfortunately not available for %s systems (Windows-only).", Sys.info()["sysname"]))
    
    hwnd <- NULL
    
  }
  
  return(hwnd)
  
}




#' Check Windows Pointer
#' 
#' Check if a pointer points to a valid and running MS-Office instance. The
#' function does this by first checking for \code{NULL} pointer
#' and then trying to get the current selection of the application.
#' 
#' @note When closing an application instance, the value of the pointer in R is
#' not somehow automatically invalidated. In such cases the corresponding
#' variable contains an invalid address.  Whether the pointer still refers to a
#' valid running application instance can be checked by
#' \code{\link{IsValidHwnd}}.

#' @aliases IsValidHwnd
#' 
#' @param hwnd the pointer to an application instance as created by 
#' \code{GetCOMAppHandle()}. 

#' @return logical value, \code{TRUE} if \code{hwnd} is a valid pointer to a running
#' application

#' @author Andri Signorell <andri@@signorell.net>

#' @seealso \code{\link{GetHwnd}()}


IsValidHwnd <- function(hwnd){
  # returns TRUE if the selection of the pointer can be evaluated
  # meaning the pointer points to a running word/excel/powerpoint instance and so far valid
  # if(!is.null(hwnd) && IsValidPtr(hwnd) )
  if(!is.null(hwnd))
    res <- !inherits(tryCatch(hwnd[["Selection"]], error=function(e) {e}), 
                     "simpleError")   # Error in
  else 
    res <- FALSE
  
  return(res)
  
}



#' Get a Handle to a Word/Excel/PowerPoint Instance
#' 
#' To be able to control MS-Office applications we need a so called handle,
#' a kind of address, to it. It is either possible to get the handle of an 
#' already running application or to start a new one.
#' 
#' The functions \code{GetNewWrd()}, \code{GetNewXL()}, and \code{GetNewPP()}
#' launch a new instance of the required application and return its handle.
#' 
#' The functions \code{GetCurrWrd()}, \code{GetCurrXL()}, and \code{GetCurrPP()}
#' look for a running instance of the specific application and return its handle.
#' Unfortunately it is not possible to choose a specific handle if there are 
#' several instances already running. The underlying RDCOM function yields 
#' some kind of "latest" launched instance. So in most cases you will be 
#' safer to start a new instance to be sure that you're communicating
#' with the expected instance. If there's no running instance the error 
#' message will be returned with a warning. 
#' 
#' @aliases GetNewWrd GetNewXL GetNewPP 
#'          GetCurrWrd GetCurrXL GetCurrPP GetHwnd

#' @param visible logical, should the application made visible? Defaults to \code{TRUE}.
#' @param newdoc logical, determining if a new document
#' should be created. Defaults to \code{TRUE}.
#' @param template the name of the template to be used for creating a new
#' Word document. Ignored by Excel and PowerPoint.

#' @return a handle (pointer) to the running Word/Excel/PowerPoint instance.

#' @note When closing an application instance, the value of the pointer in R is
#' not somehow automatically invalidated. In such cases the corresponding
#' variable contains an invalid address.  Whether the pointer still refers to a
#' valid running application instance can be checked by
#' \code{\link{IsValidHwnd}}.
#' @author Andri Signorell <andri@@signorell.net>

#' @seealso \code{\link{GetNewWrd}}, \code{\link{IsValidHwnd}}
#' 
#' @keywords misc
#' @examples
#' 
#' \dontrun{# Windows-specific example
#' 
#' # Start a new instance
#' wrd <- GetNewWrd()
#' ToWrd("Send some text to Word ... \n", wrd=wrd)
#' 
#' # Release the handle
#' rm(wrd)
#' 
#' # regrab the handle to this instance
#' wrd <- GetCurrWrd()
#' 
#' # this should be valid
#' IsValidHwnd(wrd)
#' 
#' # close the instance
#' wrd$quit()
#' 
#' # now it should be gone and the pointer invalid
#' if(IsValidHwnd(wrd)){ 
#'   print("Ouups! Still there?")
#' } else {  
#'   print("GetCurrWrd: no running word instance found...")
#' }
#' 
#' 
#' # Same with PowerPoint:
#' # get a handle to a new PowerPoint instance
#' pp <- GetNewPP()
#' 
#' # send some text
#' PpText("Some text on a slide!\n", 
#'        fontname="Arial", x=200, y=70, height=30, fontsize=14,
#'        bold=TRUE, pp=pp, bg="lemonchiffon", hasFrame=TRUE)
#' 
#' }
 



#' @rdname GetHandle
#' @export GetNewWrd
GetNewWrd <- function (visible = TRUE, newdoc = TRUE, template = "Normal") {
  
  hwnd <- GetCOMAppHandle("Word.Application", option = "lastWord", 
                          existing = FALSE, visible = TRUE)
  
  # add new document
  if (!is.null(hwnd))
    if(newdoc)
      hwnd[["Documents"]]$Add(template, FALSE, 0)
  
  # return the handle invisibly
  return(hwnd)
  
}


#' @rdname GetHandle
#' @export GetNewXL
GetNewXL <- function(visible = TRUE, newdoc = TRUE) {
  
  hwnd <- GetCOMAppHandle("Excel.Application", option="lastXL", 
                          existing=FALSE, visible=TRUE)
  
  if(!is.null(hwnd)){
    
    # Create a new workbook
    # react the same as GetNewWrd(), Word is also starting with a new document
    # whereas XL would not
    if(newdoc)      hwnd[["Workbooks"]]$Add()
    
  }
  
  return(hwnd)
  
}

#' @rdname GetHandle
#' @export GetNewPP
GetNewPP <- function (visible = TRUE) {
  
  hwnd <- GetCOMAppHandle("PowerPoint.Application", option="lastPP", 
                          existing=FALSE, visible=TRUE)
  
  if(!is.null(hwnd)){
    
    newpres <- hwnd[["Presentations"]]$Add(TRUE)
    ppLayoutBlank <- 12
    newpres[["Slides"]]$Add(1, ppLayoutBlank)
    
  }
  
  return(hwnd)  
  
}



#' @rdname GetHandle
#' @export GetCurrWrd
GetCurrWrd <- function() {
  GetCOMAppHandle("Word.Application", option="lastWord", existing=TRUE)
}


#' @rdname GetHandle
#' @export GetCurrXL
GetCurrXL <- function() {
  GetCOMAppHandle("Excel.Application", option="lastXL", existing=TRUE)
}


#' @rdname GetHandle
#' @export GetCurrPP
GetCurrPP <- function() {
  GetCOMAppHandle("PowerPoint.Application", option="lastPP", existing=TRUE)
}




#' End Application Task
#' 
#' \code{WrdKill()} and \code{XLKill()} will shut down a running application instance 
#' (which also might be invisible).
#' Background is the fact, that the simple quit() command not always
#' terminates a running XL task, and only sets it invisible (which can be 
#' observed the TaskManager). 
#' This ghost instance may sometimes confuse XLView and hinder to
#' create a new instance. In such cases we have to do the garbage
#' collection and "killing" the process seems ultima ratio.
#' 
#' @aliases WrdKill XLKill
#' 
#' @author Andri Signorell <andri@@signorell.net>
#' 
#' @seealso \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}

#' @keywords misc
#' @examples
#' \dontrun{ # Windows-specific example
#' # get a handle to a new Word instance
#' wrd <- GetNewWrd()
#' # end it with the crowbar 
#' WrdKill()
#' 
#' # get a handle to a new Excel instance
#' xl <- GetNewXL()
#' # end it with the crowbar 
#' XLKill()
#' }
#' 

#' @rdname KillApp
#' @export WrdKill
WrdKill <- function(){
  # Word might not always quit and end the task
  # so killing the task is "ultima ratio"...
  
  shell('taskkill /F /IM WINWORD.EXE')
}



#' @rdname KillApp
#' @export XLKill
XLKill <- function(){
  # Excel would only quit, when all workbooks are closed before, someone said.
  # http://stackoverflow.com/questions/15697282/excel-application-not-quitting-after-calling-quit
  
  # We experience, that it would not even then quit, when there's no workbook 
  # loaded at all.
  # maybe gc() would help ??
  # so killing the task is "ultima ratio"...
  
  shell('taskkill /F /IM EXCEL.EXE')
}





