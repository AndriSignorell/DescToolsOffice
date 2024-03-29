

#' Word VBA Constants
#' 
#' This is a list with all VBA constants for MS Word 2010, which is useful for
#' writing R functions based on recorded macros in Word. This way the constants
#' need not be replaced by their numeric values and can only be complemented
#' with the list's name, say the VBA-constant \code{wd10Percent} for example
#' can be replaced by \code{wdConst$wd10Percent}. A handful constants for Excel
#' are consolidated in \code{xlConst}.
#' 
#' 
#' @name wdConst
#' @aliases wdConst xlConst
#' @docType data
#' @format The format is:\cr List of 2755\cr $ wd100Words: num -4\cr $
#' wd10Percent: num -6\cr $ wd10Sentences: num -2\cr ...\cr
#' @source Microsoft
#' @keywords datasets

NULL



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




# Set Main Title or a header 

# if (template=="Normal" && header) 
#   .WrdPrepRep(wrd = hwnd, main = main)
# 
# # Check for existance of bookmark Main and update if found
# if(!is.null(WrdBookmark(name = "Main", wrd = hwnd))){
#   WrdUpdateBookmark(name="Main", text = main, wrd=hwnd)
#   WrdUpdateFields(wrd=hwnd, where = c(1,7))
# }





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


# Here's a list of some frequently used commands.
# Let's assume:
#   
#   xl <- GetNewXL() 
# workbooks	xl$workbooks()$count()
# quit excel	xl$quit(
# 
