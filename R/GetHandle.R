# GetHandle to Office Application


createCOMReference <- function(ref, className) {
  RDCOMClient::createCOMReference(ref, className)
}



# IsValidPtr <- function(pointer) {
#   if(is(pointer, "externalptr") | is(pointer, "COMIDispatch"))
#     !.Call("isnil", pointer)
#   else 
#     FALSE
# }


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




GetCOMAppHandle <- function(app, option=NULL, existing=FALSE, visible=NULL){
  
  if (requireNamespace("RDCOMClient", quietly = FALSE)) {
    
    if(!existing)
      # there's no "get"-function in RDCOMClient, so just create a new here..
      hwnd <- RDCOMClient::COMCreate(app, existing=existing)
    else
      hwnd <- RDCOMClient::getCOMInstance(app)
    
    if(is.null(hwnd)) 
      warning(gettext("No running %s application found!", app))
    else
      if(!is.null(visible))     hwnd[["Visible"]] <- visible
      
      
      # set the DescTools option, if required
      if(!is.null(option))
        # eval(parse(text=gettextf("options(%s = hwnd)", option)))
        eval(parse(text=gettextf("DescToolsOfficeOptions(%s = hwnd)", option)))
      
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





#' Get a Handle to a Running Word/Excel/PowerPoint Instance
#' 
#' Look for a running Word, resp. Excel instance and return its handle. If no
#' running instance is found a new instance will be created (which will be
#' communicated with a warning).
#' 
#' 
#' @aliases GetCurrWrd
#' @return a handle (pointer) to the running Word, resp. Excel instance.
#' @note When closing an application instance, the value of the pointer in R is
#' not somehow automatically invalidated. In such cases the corresponding
#' variable contains an invalid address.  Whether the pointer still refers to a
#' valid running application instance can be checked by
#' \code{\link{IsValidHwnd}}.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{GetNewWrd}}, \code{\link{IsValidHwnd}}
#' @keywords misc
#' @examples
#' 
#' \dontrun{# Windows-specific example
#' 
#' # Start a new instance
#' GetNewWrd()
#' 
#' # grab the handle to this instance
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
#' }
#' 
#' @export GetCurrWrd
GetCurrWrd <- function() {
  hwnd <- GetCOMAppHandle("Word.Application", option="lastWord", existing=TRUE)
}






#' Create a New Word Instance
#' 
#' Start a new instance of Word and return its handle. By means of this handle
#' we can then control the word application. \cr \code{WrdKill} ends a running
#' MS-Word task.
#' 
#' The package \bold{RDCOMClient} reveals the whole VBA-world of MS-Word. So
#' generally speaking any VBA code can be run fully controlled by R. In
#' practise, it might be a good idea to record a macro and rewrite the VB-code
#' in R.\cr
#' 
#' Here's a list of some frequently used commands. Let's assume we have a
#' handle to the application and a handle to the current selection defined as:
#' \preformatted{ wrd <- GetNewWrd() sel <- wrd$Selection() } Then we can
#' access the most common properties as follows: \tabular{ll}{ new document
#' \tab \code{wrd[["Documents"]]$Add(template, FALSE, 0)}, template is the
#' templatename. \cr open document \tab
#' \code{wrd[["Documents"]]$Open(Filename="C:/MyPath/MyDocument.docx")}. \cr
#' save document \tab
#' \code{wrd$ActiveDocument()$SaveAs2(FileName="P:/MyFile.docx")} \cr quit word
#' \tab \code{wrd$quit()} \cr kill word task \tab \code{WrdKill} kills a
#' running word task (which might not be ended with quit.) \cr normal text \tab
#' Use \code{\link{ToWrd}} which offers many arguments as fontname, size,
#' color, alignment etc. \cr \tab \code{ToWrd("Lorem ipsum dolor sit amet,
#' consetetur", }\cr \tab \code{font=list(name="Arial", size=10,
#' col=wdConst$wdColorRed)} \cr simple text \tab \code{sel$TypeText("sed diam
#' nonumy eirmod tempor invidunt ut labore")} \cr heading \tab
#' \code{WrdCaption("My Word-Story", index=1)} \cr insert R output \tab
#' \code{ToWrd(capture.output(str(d.diamonds)))} \cr pagebreak \tab
#' \code{sel$InsertBreak(wdConst$wdPageBreak)} \cr sectionbreak \tab
#' \code{sel$InsertBreak(wdConst$wdSectionBreakContinuous)} \cr\tab
#' (\code{wdSectionBreakNextPage}) \cr move cursor right \tab
#' \code{sel$MoveRight(Unit=wdConst$wdCharacter, Count=2,
#' Extend=wdConst$wdExtend)} \cr goto end \tab
#' \code{sel$EndKey(Unit=wdConst$wdStory)} \cr pagesetup \tab
#' \code{sel[["PageSetup"]][["Bottommargin"]] <- 4 * 72} \cr orientation \tab
#' \code{sel[["PageSetup"]][["Orientation"]] <- wdConst$wdOrientLandscape} \cr
#' add bookmark \tab
#' \code{wrd[["ActiveDocument"]][["Bookmarks"]]$Add("myBookmark")} \cr goto
#' bookmark \tab \code{sel$GoTo(wdConst$wdGoToBookmark, 0, 0, "myBookmark")}
#' \cr update bookmark \tab \code{WrdUpdateBookmark("myBookmark", "New text for
#' my bookmark")} \cr show document map \tab \code{
#' wrd[["ActiveWindow"]][["DocumentMap"]] <- TRUE} \cr create table \tab
#' \code{\link{WrdTable}}() which allows to define the table's geometry \cr
#' insert caption \tab \code{sel$InsertCaption(Label="Abbildung",
#' TitleAutoText="InsertCaption",}\cr \tab \code{Title="My Title")} \cr tables
#' of figures \tab
#' \code{wrd$ActiveDocument()$TablesOfFigures()$Add(Range=sel$range(),}\cr \tab
#' \code{Caption="Figure")} \cr insert header \tab \code{wview <-
#' wrd[["ActiveWindow"]][["ActivePane"]][["View"]][["SeekView"]] }\cr \tab
#' \code{wview <- ifelse(header, wdConst$wdSeekCurrentPageHeader,
#' wdConst$wdSeekCurrentPageFooter) }\cr \tab \code{ToWrd(x, ..., wrd=wrd) }\cr
#' 
#' }
#' 
#' @aliases GetNewWrd WrdKill createCOMReference
#' @param visible logical, should Word made visible? Defaults to \code{TRUE}.
#' @param template the name of the template to be used for creating a new
#' document.
#' @param header logical, should a caption and a list of contents be inserted?
#' Default is \code{FALSE}.
#' @param main the main title of the report
#' @return a handle (pointer) to the created Word instance.
#' @note Note that the list of contents has to be refreshed by hand after
#' inserting text (if inserted by \code{header = TRUE}).
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{GetNewXL}}, \code{\link{GetNewPP}}
#' @keywords misc
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' 
#' wrd <- GetNewWrd()
#' Desc(d.pizza[,1:4], wrd=wrd)
#' 
#' wrd <- GetNewWrd(header=TRUE)
#' Desc(d.pizza[,1:4], wrd=wrd)
#' 
#' # enumerate all bookmarks in active document
#' for(i in 1:wrd[["ActiveDocument"]][["Bookmarks"]]$count()){
#'   print(wrd[["ActiveDocument"]][["Bookmarks"]]$Item(i)$Name())
#' }
#' }
#' 
GetNewWrd <- function (visible = TRUE, template = "Normal", header = FALSE, 
                       main = "Descriptive report") {
  
  hwnd <- GetCOMAppHandle("Word.Application", option = "lastWord", 
                          existing = FALSE, visible = TRUE)
  
  if (!is.null(hwnd)) {
    newdoc <- hwnd[["Documents"]]$Add(template, FALSE, 0)
    
    if (template=="Normal" && header) 
      .WrdPrepRep(wrd = hwnd, main = main)
    
    # Check for existance of bookmark Main and update if found
    if(!is.null(WrdBookmark(name = "Main", wrd = hwnd))){
      WrdUpdateBookmark(name="Main", text = main, wrd=hwnd)
      WrdUpdateFields(wrd=hwnd, where = c(1,7))
    }
  }
  
  invisible(hwnd)
}



# wdCommentsStory = 4,
# wdEndnoteContinuationNoticeStory = 17,
# wdEndnoteContinuationSeparatorStory = 16,
# wdEndnoteSeparatorStory = 15,
# wdEndnotesStory = 3,
# wdEvenPagesFooterStory = 8,
# wdEvenPagesHeaderStory = 6,
# wdFirstPageFooterStory = 11,
# wdFirstPageHeaderStory = 10,
# wdFootnoteContinuationNoticeStory = 14,
# wdFootnoteContinuationSeparatorStory = 13,
# wdFootnoteSeparatorStory = 12,
# wdFootnotesStory = 2,
# wdMainTextStory = 1,
# wdPrimaryFooterStory = 9,
# wdPrimaryHeaderStory = 7,
# wdTextFrameStory = 5)






#' Create a New Excel Instance
#' 
#' Start a new instance of Excel and return its handle. This is needed to
#' address the Excel application and objects afterwards.
#' 
#' Here's a list of some frequently used commands.\cr Let's assume:
#' \preformatted{xl <- GetNewXL() } \tabular{ll}{ workbooks \tab
#' \code{xl$workbooks()$count()} \cr quit excel \tab \code{xl$quit()} \cr }
#' 
#' #' \code{XLKill} will kill a running XL instance (which might be invisible).
#' Background is the fact, that the simple XL$quit() command would not
#' terminate a running XL task, but only set it invisible (observe the
#' TaskManager). This ghost version may sometimes confuse XLView and hinder to
#' create a new instance. In such cases you have to do the garbage
#' collection...
#' 
#' 
#' @param visible logical, should Excel made visible? Defaults to \code{TRUE}.
#' @param newdoc logical, determining if a new workbook should be created.
#' Defaults to \code{TRUE}.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{XLView}}, \code{\link{XLGetRange}},
#' \code{\link{XLGetWorkbook}}
#' @keywords misc
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' # get a handle to a new excel instance
#' xl <- GetNewXL()
#' }
#' 
#' @export GetNewXL
GetNewXL <- function(visible = TRUE, newdoc = TRUE) {
  
  hwnd <- GetCOMAppHandle("Excel.Application", option="lastXL", existing=FALSE, visible=TRUE)
  
  if(!is.null(hwnd)){
    
    # Create a new workbook
    # react the same as GetNewWrd(), Word is also starting with a new document
    # whereas XL would not
    if(newdoc)      hwnd[["Workbooks"]]$Add()
    
  }
  
  invisible(hwnd)
  
}

#' @rdname GetNewXL
#' @export GetCurrXL
GetCurrXL <- function() {
  
  hwnd <- GetCOMAppHandle("Excel.Application", option="lastXL", existing=TRUE)
  invisible(hwnd)
}


#' @rdname GetNewXL
GetNewPP <- function (visible = TRUE) {
  
  hwnd <- GetCOMAppHandle("PowerPoint.Application", option="lastPP", existing=FALSE, visible=TRUE)
  
  if(!is.null(hwnd)){
    
    newpres <- hwnd[["Presentations"]]$Add(TRUE)
    ppLayoutBlank <- 12
    newpres[["Slides"]]$Add(1, ppLayoutBlank)
    
  }
  
  invisible(hwnd)  
  
}

#' @rdname GetNewXL
GetCurrPP <- function() {
  
  hwnd <- GetCOMAppHandle("PowerPoint.Application", option="lastPP", existing=TRUE)
  invisible(hwnd)
}



#' @rdname GetNewWrd
WrdKill <- function(){
  # Word might not always quit and end the task
  # so killing the task is "ultima ratio"...
  
  shell('taskkill /F /IM WINWORD.EXE')
}




#' @rdname GetNewXL
XLKill <- function(){
  # Excel would only quit, when all workbooks are closed before, someone said.
  # http://stackoverflow.com/questions/15697282/excel-application-not-quitting-after-calling-quit
  
  # We experience, that it would not even then quit, when there's no workbook loaded at all.
  # maybe gc() would help ??
  # so killing the task is "ultima ratio"...
  
  shell('taskkill /F /IM EXCEL.EXE')
}






