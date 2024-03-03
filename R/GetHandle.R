

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



GetCurrWrd <- function() {
  hwnd <- GetCOMAppHandle("Word.Application", option="lastWord", existing=TRUE)
}



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


GetCurrXL <- function() {
  
  hwnd <- GetCOMAppHandle("Excel.Application", option="lastXL", existing=TRUE)
  invisible(hwnd)
}



GetNewPP <- function (visible = TRUE, template = "Normal") {
  
  hwnd <- GetCOMAppHandle("PowerPoint.Application", option="lastPP", existing=FALSE, visible=TRUE)
  
  if(!is.null(hwnd)){
    
    newpres <- hwnd[["Presentations"]]$Add(TRUE)
    ppLayoutBlank <- 12
    newpres[["Slides"]]$Add(1, ppLayoutBlank)
    
  }
  
  invisible(hwnd)  
  
}


GetCurrPP <- function() {
  
  hwnd <- GetCOMAppHandle("PowerPoint.Application", option="lastPP", existing=TRUE)
  invisible(hwnd)
}

