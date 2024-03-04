.WrdPrepRep <- function(wrd, main="Bericht" ){
  
  # only internal user out from GetNewWrd()
  # creates new word instance and prepares document for report
  
  # constants
  # wdPageBreak <- 7
  # wdSeekCurrentPageHeader <- 9  ### Kopfzeile
  # wdSeekCurrentPageFooter <- 10	### Fusszeile
  # wdSeekMainDocument <- 0
  # wdPageFitBestFit <- 2
  # wdFieldEmpty <- -1
  
  # Show DocumentMap
  wrd[["ActiveWindow"]][["DocumentMap"]] <- TRUE
  wrdWind <- wrd[["ActiveWindow"]][["ActivePane"]][["View"]][["Zoom"]]
  wrdWind[["PageFit"]] <- wdConst$wdPageFitBestFit
  
  wrd[["Selection"]]$TypeParagraph()
  wrd[["Selection"]]$TypeParagraph()
  
  wrd[["Selection"]]$WholeStory()
  # 15.1.2012 auskommentiert: WrdSetFont(wrd=wrd)
  
  # Idee: ueberschrift definieren (geht aber nicht!)
  #wrd[["ActiveDocument"]][["Styles"]]$Item("ueberschrift 2")[["Font"]][["Name"]] <- "Consolas"
  #wrd[["ActiveDocument"]][["Styles"]]$Item("ueberschrift 2")[["Font"]][["Size"]] <- 10
  #wrd[["ActiveDocument"]][["Styles"]]$Item("ueberschrift 2")[["Font"]][["Bold"]] <- TRUE
  
  #wrd[["ActiveDocument"]][["Styles"]]$Item("ueberschrift 2")[["ParagraphFormat"]]["Borders"]]$Item(wdBorderTop)[["LineStyle"]] <- wdConst$wdLineStyleSingle
  
  WrdCaption( main, wrd=wrd)
  wrd[["Selection"]]$TypeText(gettextf("%s/%s\n",format(Sys.time(), "%d.%m.%Y"), Sys.getenv("username")))
  wrd[["Selection"]]$InsertBreak( wdConst$wdPageBreak)
  
  # Inhaltsverzeichnis einfuegen ***************
  wrd[["ActiveDocument"]][["TablesOfContents"]]$Add( wrd[["Selection"]][["Range"]] )
  # Original VB-Code:
  # With ActiveDocument
  # .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
  # True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
  # LowerHeadingLevel:=2, IncludePageNumbers:=True, AddedStyles:="", _
  # UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
  # True
  # .TablesOfContents(1).TabLeader = wdTabLeaderDots
  # .TablesOfContents.Format = wdIndexIndent
  # End With
  
  # Fusszeile	***************
  wrdView <- wrd[["ActiveWindow"]][["ActivePane"]][["View"]]
  wrdView[["SeekView"]] <- wdConst$wdSeekCurrentPageFooter
  wrd[["Selection"]]$TypeText( gettextf("%s/%s\t\t",format(Sys.time(), "%d.%m.%Y"), Sys.getenv("username")) )
  wrd[["Selection"]][["Fields"]]$Add( wrd[["Selection"]][["Range"]], wdConst$wdFieldEmpty, "PAGE" )
  # Roland wollte das nicht (23.11.2014):
  # wrd[["Selection"]]$TypeText("\n\n")
  wrdView[["SeekView"]] <- wdConst$wdSeekMainDocument
  
  wrd[["Selection"]]$InsertBreak( wdConst$wdPageBreak)
  invisible()
  
}




# put that to an example...
# WrdPageBreak <- function( wrd = .lastWord ) {
#   wrd[["Selection"]]$InsertBreak(wdConst$wdPageBreak)
# }





#' Insert Caption to Word
#' 
#' Insert a caption in a given level to a Word document. The caption is
#' inserted at the current cursor position.
#' 
#' 
#' @param x the text of the caption.
#' @param index integer from 1 to 9, defining the number of the heading style.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{ToWrd}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' wrd <- GetNewWrd()
#' 
#' # insert a title in level 1
#' WrdCaption("My First Caption level 1", index=1, wrd=wrd)
#' 
#' # works as well for several levels
#' sapply(1:5, function(i)
#'   WrdCaption(gettextf("My First Caption level %s",i), index=i, wrd=wrd)
#' )
#' }
#' 
#' @export WrdCaption
WrdCaption <- function(x, index = 1, wrd = DescToolsOfficeOptions("lastWord")){
  
  lst <- Recycle(x=x, index=index)
  x <-
    index <- lst[["index"]]
  for(i in seq(attr(lst, "maxdim")))
    ToWrd.character(paste(lst[["x"]][i], "\n", sep = ""),
                    style = eval(parse(text = gettextf("wdConst$wdStyleHeading%s", lst[["index"]][i]))))
  invisible()
  
}





#' Draw Borders to a Word Table
#' 
#' Drawing borders in a Word table is quite tedious. This function allows to
#' select any range and draw border lines around it.
#' 
#' 
#' @param wtab a pointer to a Word table as returned by \code{\link{WrdTable}}
#' or \code{\link{TOne}}.
#' @param from integer, a vector with two elements specifying the left upper
#' bound of the cellrange.
#' @param to integer, a vector with two elements specifying the right bottom of
#' the cellrange.
#' @param border a Word constant (\code{wdConst$wdBorder...}) defining the side
#' of the border.
#' @param lty a Word constant (\code{wdConst$wdLineStyle...}) defining the line
#' type.
#' @param col a Word constant (\code{wdConst$wdColor...}) defining the color of
#' the border. See examples for converting R colors to Word colors.
#' @param lwd a Word constant (\code{wdConst$wdLineWidth...pt}) defining the
#' line width.
#' @return nothing
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{WrdTable}}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' 
#' # create table
#' tab <- table(op=d.pizza$operator, area=d.pizza$area)
#' 
#' # send it to Word
#' wrd <- GetNewWrd()
#' wtab <- ToWrd(tab, wrd=wrd, tablestyle = NA)
#' 
#' # draw borders
#' WrdTableBorders(wtab, from=c(2,2), to=c(3,3), border=wdConst$wdBorderBottom, wrd=wrd)
#' WrdTableBorders(wtab, from=c(2,2), to=c(3,3), border=wdConst$wdBorderDiagonalUp, wrd=wrd)
#' 
#' # demonstrate linewidth and color
#' wtab <- ToWrd(tab, wrd=wrd, tablestyle = NA)
#' WrdTableBorders(wtab, col=RgbToLong(ColToRgb("olivedrab")),
#'                 lwd=wdConst$wdLineWidth150pt, wrd=wrd)
#' 
#' WrdTableBorders(wtab, border=wdConst$wdBorderBottom,
#'                 col=RgbToLong(ColToRgb("dodgerblue")),
#'                 lwd=wdConst$wdLineWidth300pt, wrd=wrd)
#' 
#' # use an R color in Word
#' RgbToLong(ColToRgb("olivedrab"))
#' 
#' # find a similar R-color for a Word color
#' ColToRgb(RgbToCol(LongToRgb(wdConst$wdColorAqua)))
#' }
#' 
#' @export WrdTableBorders
WrdTableBorders <- function (wtab, from = NULL, to = NULL, border = NULL,
                             lty = wdConst$wdLineStyleSingle, col=wdConst$wdColorBlack,
                             lwd = wdConst$wdLineWidth050pt) {
  # paint borders of a table
  
  if(is.null(from))
    from <- c(1,1)
  
  if(is.null(to))
    to <- c(wtab[["Rows"]]$Count(), wtab[["Columns"]]$Count())
  
  wrd <- wtab[["Application"]]
  rng <- wrd[["ActiveDocument"]]$Range(start=wtab$Cell(from[1], from[2])[["Range"]][["Start"]],
                                       end=wtab$Cell(to[1], to[2])[["Range"]][["End"]])
  
  rng$Select()
  
  if(is.null(border))
    # use all borders by default
    border <- wdConst[c("wdBorderTop","wdBorderBottom","wdBorderLeft","wdBorderRight",
                        "wdBorderHorizontal","wdBorderVertical")]
  
  for(b in border){
    wborder <- wrd[["Selection"]]$Borders(b)
    wborder[["LineStyle"]] <- lty
    wborder[["Color"]] <- col
    wborder[["LineWidth"]] <- lwd
  }
  
  invisible()
}








#' Return the Cell Range Of a Word Table
#' 
#' Return a handle of a cell range of a word table. This is useful for
#' formating the cell range.
#' 
#' Cell range selecting might be complicated. This function makes it easy.
#' 
#' @param wtab a handle to the word table as returned i.g. by
#' \code{\link{WrdTable}}
#' @param from a vector containing row- and column number of the left/upper
#' cell of the cell range.
#' @param to a vector containing row- and column number of the right/lower cell
#' of the cell range.
#' @return a handle to the range.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{WrdTable}}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' 
#' # Windows-specific example
#' wrd <- GetNewWrd()
#' WrdTable(nrow=3, ncol=3, wrd=wrd)
#' crng <- WrdCellRange(from=c(1,2), to=c(2,3))
#' crng$Select()
#' }
#' 
#' @export WrdCellRange
WrdCellRange <- function(wtab, from, to) {
  # returns a handle for the table range
  wtrange <- wtab[["Parent"]]$Range(
    wtab$Cell(from[1], from[2])[["Range"]][["Start"]],
    wtab$Cell(to[1], to[2])[["Range"]][["End"]]
  )
  
  return(wtrange)
}




#' Merges Cells Of a Defined Word Table Range
#' 
#' Merges a cell range of a word table.
#' 
#' 
#' @param wtab a handle to the word table as returned i.g. by
#' \code{\link{WrdTable}}
#' @param rstart the left/upper cell of the cell range.
#' @param rend the right/lower cell of the cell range.
#' @return nothing
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{WrdTable}}, \code{\link{WrdCellRange}}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' 
#' # Windows-specific example
#' wrd <- GetNewWrd()
#' wtab <- WrdTable(nrow=3, ncol=3, wrd=wrd)
#' WrdMergeCells(wtab, rstart=c(1,2), rend=c(2,3))
#' }
#' 
#' @export WrdMergeCells
WrdMergeCells <- function(wtab, rstart, rend) {
  
  rng <- WrdCellRange(wtab, rstart, rend)
  rng[["Cells"]]$Merge()
  
}



#' Format Cells Of a Word Table
#' 
#' Format cells of a Word table.
#' 
#' Cell range selecting might be complicated. This function makes it easy.
#' 
#' @param wtab a handle to the word table as returned i.g. by
#' \code{\link{WrdTable}}
#' @param rstart the left/upper cell of the cell range
#' @param rend the right/lower cell of the cell range
#' @param col the foreground colour
#' @param bg the background colour
#' @param font the font to be used to the output. This should be defined as a
#' list containing fontname, fontsize, bold and italic flags:\cr
#' \code{list(name="Arial", size=10, bold=FALSE, italic=TRUE,
#' color=wdConst$wdColorBlack)}.
#' @param border the border of the cell range, defined as a list containing
#' arguments for border, linestyle, linewidth and color. \code{border} is a
#' vector containing the parts of the border defined by the Word constants
#' \code{wdConst$wdBorder...}, being $wdBorderBottom, $wdBorderLeft,
#' $wdBorderTop, $wdBorderRight, $wdBorderHorizontal, $wdBorderVertical,
#' $wdBorderDiagonalUp, $wdBorderDiagonalDown. linestyle, linewidth and color
#' will be recycled to the required dimension.
#' @param align a character out of \code{"l"}, \code{"c"}, \code{"r"} setting
#' the horizontal alignment of the cell range.
#' @return a handle to the range.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{WrdTable}}
#' @keywords print
#' @examples
#' 
#' \dontrun{   # Windows-specific example
#' 
#' m <- matrix(rnorm(12)*100, nrow=4,
#'             dimnames=list(LETTERS[1:4], c("Variable","Value","Remark")))
#' 
#' wrd <- GetNewWrd()
#' wt <- ToWrd(m)
#' 
#' WrdFormatCells(wt, rstart=c(3,1), rend=c(4,3),
#'                bg=wdConst$wdColorGold, font=list(name="Arial Narrow", bold=TRUE),
#'                align="c", border=list(color=wdConst$wdColorTeal,
#'                                       linewidth=wdConst$wdLineWidth300pt))
#' 
#' }
#' 
#' @export WrdFormatCells
WrdFormatCells <- function(wtab, rstart, rend, col=NULL, bg=NULL, font=NULL,
                           border=NULL, align=NULL){
  
  
  rng <- WrdCellRange(wtab, rstart, rend)
  shad <- rng[["Shading"]]
  
  if (!is.null(col))
    shad[["ForegroundPatternColor"]] <- col
  
  if (!is.null(bg))
    shad[["BackgroundPatternColor"]] <- bg
  
  wrdFont <- rng[["Font"]]
  if (!is.null(font$name))
    wrdFont[["Name"]] <- font$name
  if (!is.null(font$size))
    wrdFont[["Size"]] <- font$size
  if (!is.null(font$bold))
    wrdFont[["Bold"]] <- font$bold
  if (!is.null(font$italic))
    wrdFont[["Italic"]] <- font$italic
  if (!is.null(font$color))
    wrdFont[["Color"]] <- font$color
  
  if (!is.null(align)) {
    align <- match.arg(align, choices = c("l", "c", "r"))
    align <- unlist(wdConst[c("wdAlignParagraphLeft",
                              "wdAlignParagraphCenter",
                              "wdAlignParagraphRight")])[match(x=align, table= c("l", "c", "r"))]
    
    rng[["ParagraphFormat"]][["Alignment"]] <- align
  }
  
  if(!is.null(border)) {
    if(identical(border, TRUE))
      # set default values
      border <- list(border=c(wdConst$wdBorderBottom,
                              wdConst$wdBorderLeft,
                              wdConst$wdBorderTop,
                              wdConst$wdBorderRight),
                     linestyle=wdConst$wdLineStyleSingle,
                     linewidth=wdConst$wdLineWidth025pt,
                     color=wdConst$wdColorBlack)
    
    if(is.null(border$border))
      border$border <- c(wdConst$wdBorderBottom,
                         wdConst$wdBorderLeft,
                         wdConst$wdBorderTop,
                         wdConst$wdBorderRight)
    
    if(is.null(border$linestyle))
      border$linestyle <- wdConst$wdLineStyleSingle
    
    border <- do.call(Recycle, border)
    
    for(i in 1:attr(border, which = "maxdim")) {
      b <- rng[["Borders"]]$Item(border$border[i])
      
      if(!is.null(border$linestyle[i]))
        b[["LineStyle"]] <- border$linestyle[i]
      
      if(!is.null(border$linewidth[i]))
        b[["LineWidth"]] <- border$linewidth[i]
      
      if(!is.null(border$color))
        b[["Color"]] <- border$color[i]
    }
  }
  
}





# Get and set font



#' Get or Set the Font in Word
#' 
#' \code{WrdFont} can be used to get and set the font in Word for the text to
#' be inserted. \code{WrdFont} returns the font at the current cursor position.
#' 
#' The font color can be defined by a Word constant beginning with
#' \code{wdConst$wdColor}. The defined colors can be listed with
#' \code{grep("wdColor", names(wdConst), val=TRUE)}.
#' 
#' @aliases WrdFont WrdFont<-
#' @param value the font to be used to the output. This should be defined as a
#' list containing fontname, fontsize, bold and italic flags:\cr
#' \code{list(name="Arial", size=10, bold=FALSE, italic=TRUE,
#' color=wdConst$wdColorBlack)}.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return a list of the attributes of the font in the current cursor position:
#' \item{name}{the fontname} \item{size}{the fontsize} \item{bold}{bold}
#' \item{italic}{italic} \item{color}{the fontcolor}
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{ToWrd}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' 
#' wrd <- GetNewWrd()
#' 
#' for(i in seq(10, 24, 2))
#'   ToWrd(gettextf("This is Arial size %s \n", i), font=list(name="Arial", size=i))
#' 
#' for(i in seq(10, 24, 2))
#'   ToWrd(gettextf("This is Times size %s \n", i), font=list(name="Times", size=i))
#' }
#' @export WrdFont
WrdFont <- function(wrd = DescToolsOfficeOptions("lastWord") ) {
  # returns the font object list: list(name, size, bold, italic) on the current position
  
  wrdSel <- wrd[["Selection"]]
  wrdFont <- wrdSel[["Font"]]
  
  currfont <- list(
    name = wrdFont[["Name"]] ,
    size = wrdFont[["Size"]] ,
    bold = wrdFont[["Bold"]] ,
    italic = wrdFont[["Italic"]],
    color = setNames(wrdFont[["Color"]], names(which(
      wdConst==wrdFont[["Color"]] & grepl("wdColor", names(wdConst)))))
  )
  
  class(currfont) <- "font"
  return(currfont)
}

#' @rdname WrdFont
`WrdFont<-` <- function(wrd, value){
  
  wrdSel <- wrd[["Selection"]]
  wrdFont <- wrdSel[["Font"]]
  
  # set the new font
  if(!is.null(value$name)) wrdFont[["Name"]] <- value$name
  if(!is.null(value$size)) wrdFont[["Size"]] <- value$size
  if(!is.null(value$bold)) wrdFont[["Bold"]] <- value$bold
  if(!is.null(value$italic)) wrdFont[["Italic"]] <- value$italic
  if(!is.null(value$color)) wrdFont[["Color"]] <- value$color
  
  return(wrd)
}



# Get and set ParagraphFormat



#' Get or Set the Paragraph Format in Word
#' 
#' \code{WrdParagraphFormat} can be used to get and set the font in Word for
#' the text to be inserted.
#' 
#' 
#' @aliases WrdParagraphFormat WrdParagraphFormat<-
#' @param value a list defining the paragraph format.  This can contain any
#' combination of: \code{LeftIndent}, \code{RightIndent}, \code{SpaceBefore},
#' \code{SpaceBeforeAuto}, \code{SpaceAfter}, \code{SpaceAfterAuto},
#' \code{LineSpacingRule}, \code{Alignment}, \code{WidowControl},
#' \code{KeepWithNext}, \code{KeepTogether}, \code{PageBreakBefore},
#' \code{NoLineNumber}, \code{Hyphenation}, \code{FirstLineIndent},
#' \code{OutlineLevel}, \code{CharacterUnitLeftIndent},
#' \code{CharacterUnitRightIndent}, \code{CharacterUnitFirstLineIndent},
#' \code{LineUnitBefore}, \code{LineUnitAfter} and/or \code{MirrorIndents}.
#' The possible values of the arguments are found in the Word constants with
#' the respective name. \cr The alignment for example can be set to
#' \code{wdAlignParagraphLeft}, \code{wdAlignParagraphRight},
#' \code{wdAlignParagraphCenter} and so on.  \cr Left alignment with
#' indentation would be set as:\cr
#' \code{list(Alignment=wdConst$wdAlignParagraphLeft, LeftIndent=42.55)}.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return an object with the class \code{paragraph}, basically a list with the
#' attributes of the paragraph in the current cursor position:
#' \item{LeftIndent}{left indentation in (in points) for the specified
#' paragraphs.} \item{RightIndent}{right indent (in points) for the specified
#' paragraphs.} \item{SpaceBefore}{spacing (in points) before the specified
#' paragraphs.} \item{SpaceBeforeAuto}{\code{TRUE} if Microsoft Word
#' automatically sets the amount of spacing before the specified paragraphs.}
#' \item{SpaceAfter}{amount of spacing (in points) after the specified
#' paragraph or text column.} \item{SpaceAfterAuto}{\code{TRUE} if Microsoft
#' Word automatically sets the amount of spacing after the specified
#' paragraphs.} \item{LineSpacingRule}{line spacing for the specified paragraph
#' formatting. Use \code{wdLineSpaceSingle}, \code{wdLineSpace1pt5}, or
#' \code{wdLineSpaceDouble} to set the line spacing to one of these values. To
#' set the line spacing to an exact number of points or to a multiple number of
#' lines, you must also set the \code{LineSpacing} property.}
#' \item{Alignment}{\code{WdParagraphAlignment} constant that represents the
#' alignment for the specified paragraphs.} \item{WidowControl}{\code{TRUE} if
#' the first and last lines in the specified paragraph remain on the same page
#' as the rest of the paragraph when Word repaginates the document. Can be
#' \code{TRUE}, \code{FALSE} or \code{wdUndefined}.}
#' \item{KeepWithNext}{\code{TRUE} if the specified paragraph remains on the
#' same page as the paragraph that follows it when Microsoft Word repaginates
#' the document. Read/write Long.} \item{KeepTogether}{\code{TRUE} if all lines
#' in the specified paragraphs remain on the same page when Microsoft Word
#' repaginates the document.} \item{PageBreakBefore}{\code{TRUE} if a page
#' break is forced before the specified paragraphs. Can be \code{TRUE},
#' \code{FALSE}, or \code{wdUndefined}.} \item{NoLineNumber}{\code{TRUE} if
#' line numbers are repressed for the specified paragraphs. Can be \code{TRUE},
#' \code{FALSE}, or \code{wdUndefined}.} \item{Hyphenation}{\code{TRUE} if the
#' specified paragraphs are included in automatic hyphenation. \code{FALSE} if
#' the specified paragraphs are to be excluded from automatic hyphenation.}
#' \item{FirstLineIndent}{value (in points) for a first line or hanging indent.
#' Use a positive value to set a first-line indent, and use a negative value to
#' set a hanging indent.} \item{OutlineLevel}{outline level for the specified
#' paragraphs.} \item{CharacterUnitLeftIndent}{left indent value (in
#' characters) for the specified paragraphs.}
#' \item{CharacterUnitRightIndent}{right indent value (in characters) for the
#' specified paragraphs. } \item{LineUnitBefore}{amount of spacing (in
#' gridlines) before the specified paragraphs. } \item{LineUnitAfter}{amount of
#' spacing (in gridlines) after the specified paragraphs.}
#' \item{MirrorIndents}{Long that represents whether left and right indents are
#' the same width. Can be \code{TRUE}, \code{FALSE}, or \code{wdUndefined}.}
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{ToWrd}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' # Windows-specific example
#' wrd <- GetNewWrd()  # get the handle to a new word instance
#' 
#' WrdParagraphFormat(wrd=wrd) <- list(Alignment=wdConst$wdAlignParagraphLeft,
#'                                     LeftIndent=42.55)
#' 
#' ToWrd("Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy
#' eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.
#' At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd
#' gubergren, no sea takimata sanctus est.\n", wrd=wrd)
#' 
#' # reset
#' WrdParagraphFormat(wrd=wrd) <- list(LeftIndent=0)
#' }
#' 
#' @export WrdParagraphFormat
WrdParagraphFormat <- function(wrd = DescToolsOfficeOptions("lastWord") ) {
  
  wrdPar <- wrd[["Selection"]][["ParagraphFormat"]]
  
  currpar <- list(
    LeftIndent               =wrdPar[["LeftIndent"]] ,
    RightIndent              =wrdPar[["RightIndent"]] ,
    SpaceBefore              =wrdPar[["SpaceBefore"]] ,
    SpaceBeforeAuto          =wrdPar[["SpaceBeforeAuto"]] ,
    SpaceAfter               =wrdPar[["SpaceAfter"]] ,
    SpaceAfterAuto           =wrdPar[["SpaceAfterAuto"]] ,
    LineSpacingRule          =wrdPar[["LineSpacingRule"]],
    Alignment                =wrdPar[["Alignment"]],
    WidowControl             =wrdPar[["WidowControl"]],
    KeepWithNext             =wrdPar[["KeepWithNext"]],
    KeepTogether             =wrdPar[["KeepTogether"]],
    PageBreakBefore          =wrdPar[["PageBreakBefore"]],
    NoLineNumber             =wrdPar[["NoLineNumber"]],
    Hyphenation              =wrdPar[["Hyphenation"]],
    FirstLineIndent          =wrdPar[["FirstLineIndent"]],
    OutlineLevel             =wrdPar[["OutlineLevel"]],
    CharacterUnitLeftIndent  =wrdPar[["CharacterUnitLeftIndent"]],
    CharacterUnitRightIndent =wrdPar[["CharacterUnitRightIndent"]],
    CharacterUnitFirstLineIndent=wrdPar[["CharacterUnitFirstLineIndent"]],
    LineUnitBefore           =wrdPar[["LineUnitBefore"]],
    LineUnitAfter            =wrdPar[["LineUnitAfter"]],
    MirrorIndents            =wrdPar[["MirrorIndents"]]
    # wrdPar[["TextboxTightWrap"]] <- TextboxTightWrap
  )
  
  class(currpar) <- "paragraph"
  return(currpar)
}


#' @rdname WrdParagraphFormat
`WrdParagraphFormat<-` <- function(wrd, value){
  
  wrdPar <- wrd[["Selection"]][["ParagraphFormat"]]
  
  # set the new font
  if(!is.null(value$LeftIndent)) wrdPar[["LeftIndent"]] <- value$LeftIndent
  if(!is.null(value$RightIndent)) wrdPar[["RightIndent"]] <- value$RightIndent
  if(!is.null(value$SpaceBefore)) wrdPar[["SpaceBefore"]] <- value$SpaceBefore
  if(!is.null(value$SpaceBeforeAuto)) wrdPar[["SpaceBeforeAuto"]] <- value$SpaceBeforeAuto
  if(!is.null(value$SpaceAfter)) wrdPar[["SpaceAfter"]] <- value$SpaceAfter
  if(!is.null(value$SpaceAfterAuto)) wrdPar[["SpaceAfterAuto"]] <- value$SpaceAfterAuto
  if(!is.null(value$LineSpacingRule)) wrdPar[["LineSpacingRule"]] <- value$LineSpacingRule
  if(!is.null(value$Alignment)) {
    if(is.character(value$Alignment))
      switch(match.arg(value$Alignment, choices = c("left","center","right"))
             , left=value$Alignment <- wdConst$wdAlignParagraphLeft
             , center=value$Alignment <- wdConst$wdAlignParagraphCenter
             , right=value$Alignment <- wdConst$wdAlignParagraphRight
      )
    wrdPar[["Alignment"]] <- value$Alignment
  }
  if(!is.null(value$WidowControl)) wrdPar[["WidowControl"]] <- value$WidowControl
  if(!is.null(value$KeepWithNext)) wrdPar[["KeepWithNext"]] <- value$KeepWithNext
  if(!is.null(value$KeepTogether)) wrdPar[["KeepTogether"]] <- value$KeepTogether
  if(!is.null(value$PageBreakBefore)) wrdPar[["PageBreakBefore"]] <- value$PageBreakBefore
  if(!is.null(value$NoLineNumber)) wrdPar[["NoLineNumber"]] <- value$NoLineNumber
  if(!is.null(value$Hyphenation)) wrdPar[["Hyphenation"]] <- value$Hyphenation
  if(!is.null(value$FirstLineIndent)) wrdPar[["FirstLineIndent"]] <- value$FirstLineIndent
  if(!is.null(value$OutlineLevel)) wrdPar[["OutlineLevel"]] <- value$OutlineLevel
  if(!is.null(value$CharacterUnitLeftIndent)) wrdPar[["CharacterUnitLeftIndent"]] <- value$CharacterUnitLeftIndent
  if(!is.null(value$CharacterUnitRightIndent)) wrdPar[["CharacterUnitRightIndent"]] <- value$CharacterUnitRightIndent
  if(!is.null(value$CharacterUnitFirstLineIndent)) wrdPar[["CharacterUnitFirstLineIndent"]] <- value$CharacterUnitFirstLineIndent
  if(!is.null(value$LineUnitBefore)) wrdPar[["LineUnitBefore"]] <- value$LineUnitBefore
  if(!is.null(value$LineUnitAfter)) wrdPar[["LineUnitAfter"]] <- value$LineUnitAfter
  if(!is.null(value$MirrorIndents)) wrdPar[["MirrorIndents"]] <- value$MirrorIndents
  
  return(wrd)
  
}




#' Get or Set the Style in Word
#' 
#' \code{WrdStyle} can be used to get and set the style in Word for the text to
#' be inserted. \code{WrdStyle} returns the style at the current cursor
#' position.
#' 
#' 
#' @aliases WrdStyle WrdStyle<-
#' @param value the name of the style to be used to the output. This should be
#' defined an existing name.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return character, name of the style
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{ToWrd}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' 
#' wrd <- GetNewWrd()
#' # the current stlye
#' WrdStyle(wrd)
#' }
#' @export WrdStyle
WrdStyle <- function (wrd = DescToolsOfficeOptions("lastWord")) {
  wrdSel <- wrd[["Selection"]]
  wrdStyle <- wrdSel[["Style"]][["NameLocal"]]
  return(wrdStyle)
}

#' @rdname WrdStyle
`WrdStyle<-` <- function (wrd, value) {
  wrdSel <- wrd[["Selection"]][["Paragraphs"]]
  wrdSel[["Style"]] <- value
  return(wrd)
}








#' Insert a Page Break
#' 
#' Insert a page break in a MS-Word (R) document at the position of the cursor.
#' 
#' 
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{WrdFont}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' wrd <- GetNewWrd()
#' WrdText("This is text on page 1.\n\n")
#' WrdPageBreak()
#' WrdText("This is text on another page.\n\n")
#' }
#' @export WrdPageBreak
WrdPageBreak <- function(wrd = DescToolsOfficeOptions("lastWord")) {
  wrd[["Selection"]]$InsertBreak(wdConst$wdSectionBreakNextPage)
  invisible()
}







WrdUpdateFields <- function(where = "wholestory", wrd = DescToolsOfficeOptions("lastWord")) {
  
  ii <- if( identical(where, "wholestory") )
    list(
      wdCommentsStory = 4,
      wdEndnoteContinuationNoticeStory = 17,
      wdEndnoteContinuationSeparatorStory = 16,
      wdEndnoteSeparatorStory = 15,
      wdEndnotesStory = 3,
      wdEvenPagesFooterStory = 8,
      wdEvenPagesHeaderStory = 6,
      wdFirstPageFooterStory = 11,
      wdFirstPageHeaderStory = 10,
      wdFootnoteContinuationNoticeStory = 14,
      wdFootnoteContinuationSeparatorStory = 13,
      wdFootnoteSeparatorStory = 12,
      wdFootnotesStory = 2,
      wdMainTextStory = 1,
      wdPrimaryFooterStory = 9,
      wdPrimaryHeaderStory = 7,
      wdTextFrameStory = 5)
  
  else
    where
  
  doc <- wrd$activedocument()
  for(i in ii) {
    
    # we cannot simply loop over a sequence 1:count() as indexing a nonexisting story raises a COMError
    # and the index of the story is not an ascending integer, but a wdStory constant
    # not found a handle to get a list of existing storyranges
    StoryRange <- tryCatch(doc$StoryRanges()[[i]], error = function(e) NULL)
    if(!is.null(StoryRange)) {
      if(StoryRange$Fields()$Count() > 0) {
        for(j in seq(StoryRange$Fields()$Count())){
          StoryRange$Fields(j)$Update()
        }
      }
    }
  }
}





WrdOpenFile <- function(fn, wrd = DescToolsOfficeOptions("lastWord")){
  
  if(!IsValidHwnd(wrd)){
    wrd <- GetNewWrd()
    wrd[["ActiveDocument"]]$Close()
  }
  
  # ChangeFileOpenDirectory "C:\Users\HK1S0\Desktop\"
  # 
  # Documents.Open FileName:="DynWord.docx", ConfirmConversions:=False, _
  #         ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
  #         PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
  #         WritePasswordTemplate:="", Format:=wdOpenFormatAuto, XMLTransform:=""
  
  res <- wrd[["Documents"]]$Open(FileName=fn)
  
  # return document
  invisible(res)
}





#' Open and Save Word Documents
#' 
#' Open and save MS-Word documents.
#' 
#' 
#' @aliases WrdSaveAs WrdOpenFile
#' @param fn filename and -path for the document.
#' @param fileformat file format, one out of \code{"doc"}, \code{"htm"},
#' \code{"pdf"}.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return nothing returned
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{GetNewWrd}()}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' #   Windows-specific example
#' wrd <- GetNewWrd()
#' WrdCaption("A Report")
#' WrdSaveAs(fn="report", fileformat="htm")
#' }
#' @export WrdSaveAs
WrdSaveAs <- function(fn, fileformat="docx", wrd = DescToolsOfficeOptions("lastWord")) {
  
  wdConst$wdExportFormatPDF <- 17
  
  if(fileformat %in% c("doc","docx"))
    wrd$ActiveDocument()$SaveAs(FileName=fn, FileFormat=wdConst$wdFormatDocument)
  else if(fileformat %in% c("htm", "html"))
    wrd$ActiveDocument()$SaveAs2(FileName=fn, FileFormat=wdConst$wdFormatHTML)
  else if(fileformat == "pdf")
    wrd$ActiveDocument()$ExportAsFixedFormat(OutputFileName="Einkommen2.pdf",
                                             ExportFormat=wdConst$wdExportFormatPDF)
  
  # ChangeFileOpenDirectory "C:\Users\HK1S0\Desktop\"
  # ActiveDocument.SaveAs2 FileName:="Einkommen.htm", FileFormat:=wdFormatHTML _
  #     , LockComments:=False, Password:="", AddToRecentFiles:=True, _
  #     WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
  #      SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
  #     False, CompatibilityMode:=0
  # ActiveWindow.View.Type = wdWebView
  #
  # ActiveDocument.ExportAsFixedFormat OutputFileName:= _
  #     "C:\Users\HK1S0\Desktop\Einkommen.pdf", ExportFormat:=wdExportFormatPDF, _
  #     OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
  #     wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
  #     IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
  #     wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
  #     True, UseISO19005_1:=False
  
  invisible()
  
}


# Example: WrdPlot(picscale=30)
#          WrdPlot(width=8)


CmToPts <- function(x) x * 28.35
PtsToCm <- function(x) x / 28.35
# http://msdn.microsoft.com/en-us/library/bb214076(v=office.12).aspx




#' Insert Active Plot to Word
#' 
#' This function inserts the plot on the active plot device to Word. The image
#' is transferred by saving the picture to a file in R and inserting the file
#' in Word. The format of the plot can be selected, as well as crop options and
#' the size factor for inserting.
#' 
#' 
#' @param type the format for the picture file, default is \code{"png"}.
#' @param append.cr should a carriage return be appended? Default is TRUE.
#' @param crop crop options for the picture, defined by a 4-elements-vector.
#' The first element is the bottom side, the second the left and so on.
#' @param main a caption for the plot. This will be inserted by InserCaption in
#' Word. Default is NULL, which will insert nothing.
#' @param picscale scale factor of the picture in percent, default ist 100.
#' @param height height in cm, this overrides the picscale if both are given.
#' @param width width in cm, this overrides the picscale if both are given.
#' @param res resolution for the png file, defaults to 300.
#' @param dfact the size factor for the graphic.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return Returns a pointer to the inserted picture.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{ToWrd}}, \code{\link{WrdCaption}},
#' \code{\link{GetNewWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' # let's have some graphics
#' plot(1,type="n", axes=FALSE, xlab="", ylab="", xlim=c(0,1), ylim=c(0,1), asp=1)
#' rect(0,0,1,1,col="black")
#' segments(x0=0.5, y0=seq(0.632,0.67, length.out=100),
#'   y1=seq(0.5,0.6, length.out=100), x1=1, col=rev(rainbow(100)))
#' polygon(x=c(0.35,0.65,0.5), y=c(0.5,0.5,0.75), border="white",
#'   col="black", lwd=2)
#' segments(x0=0,y0=0.52, x1=0.43, y1=0.64, col="white", lwd=2)
#' x1 <- seq(0.549,0.578, length.out=50)
#' segments(x0=0.43, y0=0.64, x1=x1, y1=-tan(pi/3)* x1 + tan(pi/3) * 0.93,
#'   col=rgb(1,1,1,0.35))
#' 
#' 
#' # get a handle to a new word instance
#' wrd <- GetNewWrd()
#' # insert plot with a specified height
#' WrdPlot(wrd=wrd, height=5)
#' ToWrd("Remember?\n", fontname="Arial", fontsize=14, bold=TRUE, wrd=wrd)
#' # crop the picture
#' WrdPlot(wrd=wrd, height=5, crop=c(9,9,0,0))
#' 
#' 
#' wpic <- WrdPlot(wrd=wrd, height=5, crop=c(9,9,0,0))
#' wpic
#' }
#' 
#' @export WrdPlot
WrdPlot <- function( type="png", append.cr=TRUE, crop=c(0,0,0,0), main = NULL,
                     picscale=100, height=NA, width=NA, res=300, dfact=1.6, wrd = DescToolsOfficeOptions("lastWord") ){
  
  # png is considered a good default choice for export to word (Smith)
  # http://blog.revolutionanalytics.com/2009/01/10-tips-for-making-your-r-graphics-look-their-best.html
  
  # height, width in cm!
  # scale will be overidden, if height/width defined
  
  
  
  # handle missing height or width values
  if (is.na(width) ){
    if (is.na(height)) {
      width <- 14
      height <- par("pin")[2] / par("pin")[1] * width
    } else {
      width <- par("pin")[1] / par("pin")[2] * height
    }
  } else {
    if (is.na(height) ){
      height <- par("pin")[2] / par("pin")[1] * width
    }
  }
  
  
  # get a [type] tempfilename:
  fn <- paste( tempfile(pattern = "file", tmpdir = tempdir()), ".", type, sep="" )
  # this is a problem for RStudio....
  # savePlot( fn, type=type )
  # png(fn, width=width, height=height, units="cm", res=300 )
  dev.copy(eval(parse(text=type)), fn, width=width*dfact, height=height*dfact, res=res, units="cm")
  d <- dev.off()
  
  # add it to our word report
  res <- wrd[["Selection"]][["InlineShapes"]]$AddPicture( fn, FALSE, TRUE )
  wrdDoc <- wrd[["ActiveDocument"]]
  pic <- wrdDoc[["InlineShapes"]]$Item( wrdDoc[["InlineShapes"]][["Count"]] )
  
  pic[["LockAspectRatio"]] <- -1  # = msoTrue
  picfrmt <- pic[["PictureFormat"]]
  picfrmt[["CropBottom"]] <- CmToPts(crop[1])
  picfrmt[["CropLeft"]] <- CmToPts(crop[2])
  picfrmt[["CropTop"]] <- CmToPts(crop[3])
  picfrmt[["CropRight"]] <- CmToPts(crop[4])
  
  if( is.na(height) & is.na(width) ){
    # or use the ScaleHeight/ScaleWidth attributes:
    pic[["ScaleHeight"]] <- picscale
    pic[["ScaleWidth"]] <- picscale
  } else {
    # Set new height:
    if( is.na(width) ) width <- height / PtsToCm( pic[["Height"]] ) * PtsToCm( pic[["Width"]] )
    if( is.na(height) ) height <- width / PtsToCm( pic[["Width"]] ) * PtsToCm( pic[["Height"]] )
    pic[["Height"]] <- CmToPts(height)
    pic[["Width"]] <- CmToPts(width)
  }
  
  if( append.cr == TRUE ) { wrd[["Selection"]]$TypeText("\n")
  } else {
    wrd[["Selection"]]$MoveRight(wdConst$wdCharacter, 1, 0)
  }
  
  if( file.exists(fn) ) { file.remove(fn) }
  
  if(!is.null(main)){
    # insert caption
    sel <- wrd$Selection()  # "Abbildung"
    sel$InsertCaption(Label=wdConst$wdCaptionFigure, Title=main)
    sel$TypeParagraph()
  }
  
  invisible(pic)
  
}





#' Insert a Table in a Word Document
#' 
#' Create a table with a specified number of rows and columns in a Word
#' document at the current position of the cursor.
#' 
#' 
#' @param nrow number of rows.
#' @param ncol number of columns.
#' @param heights a vector of the row heights (in [cm]). If set to \code{NULL}
#' (which is the default) the Word defaults will be used. The values will be
#' recyled, if necessary.
#' @param widths a vector of the column widths (in [cm]). If set to \code{NULL}
#' (which is the default) the Word defaults will be used. The values will be
#' recyled, if necessary.
#' @param main a caption for the plot. This will be inserted by InserCaption in
#' Word. Default is NULL, which will insert nothing.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return A pointer to the inserted table.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{GetNewWrd}}, \code{\link{ToWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{ # Windows-specific example
#' wrd <- GetNewWrd()
#' WrdTable(nrow=3, ncol=3, wrd=wrd)
#' }
#' 
#' @export WrdTable
WrdTable <- function(nrow = 1, ncol = 1, heights = NULL, widths = NULL, main = NULL, wrd = DescToolsOfficeOptions("lastWord")){
  
  res <- wrd[["ActiveDocument"]][["Tables"]]$Add(wrd[["Selection"]][["Range"]],
                                                 NumRows = nrow, NumColumns = ncol)
  if(!is.null(widths)) {
    widths <- rep(widths, length.out=ncol)
    for(i in 1:ncol){
      # set column-widths
      tcol <- res$Columns(i)
      tcol[["Width"]] <- CmToPts(widths[i])
    }
  }
  if(!is.null(heights)) {
    heights <- rep(heights, length.out=nrow)
    for(i in 1:nrow){
      # set row heights
      tcol <- res$Rows(i)
      tcol[["Height"]] <- CmToPts(heights[i])
    }
  }
  
  if(!is.null(main)){
    # insert caption
    sel <- wrd$Selection()  # "Abbildung"
    sel$InsertCaption(Label=wdConst$wdCaptionTable, Title=main)
    sel$TypeParagraph()
  }
  
  invisible(res)
}




###

# ## Word Table - experimental code
#
# WrdTable <- function(tab, main = NULL, wrd = DescToolsOfficeOptions("lastWord"), row.names = FALSE, ...){
#   UseMethod("WrdTable")
#
# }
#
#
# WrdTable.Freq <- function(tab, main = NULL, wrd = DescToolsOfficeOptions("lastWord"), row.names = FALSE, ...){
#
#   tab[,c(3,5)] <- sapply(round(tab[,c(3,5)], 3), Format, digits=3)
#   res <- WrdTable.default(tab=tab, wrd=wrd)
#
#   if(!is.null(main)){
#     # insert caption
#     sel <- wrd$Selection()  # "Abbildung"
#     sel$InsertCaption(Label=wdConst$wdCaptionTable, Title=main)
#     sel$TypeParagraph()
#   }
#
#   invisible(res)
#
# }
#
# WrdTable.ftable <- function(tab, main = NULL, wrd = DescToolsOfficeOptions("lastWord"), row.names = FALSE, ...) {
#   tab <- FixToTable(capture.output(tab))
#   NextMethod()
# }
#
#
# WrdTable.default <- function (tab, font = NULL, align=NULL, autofit = TRUE, main = NULL,
#                               wrd = DescToolsOfficeOptions("lastWord"), row.names=FALSE,
#                               ...) {
#
#   dim1 <- ncol(tab)
#   dim2 <- nrow(tab)
#   if(row.names) dim1 <- dim1 + 1
#
#   # wdConst ist ein R-Objekt (Liste mit 2755 Objekten!!!)
#
#   write.table(tab, file = "clipboard", sep = "\t", quote = FALSE, row.names=row.names)
#
#   myRange <- wrd[["Selection"]][["Range"]]
#   bm      <- wrd[["ActiveDocument"]][["Bookmarks"]]$Add("PasteHere", myRange)
#   myRange$Paste()
#
#   if(row.names) wrd[["Selection"]]$TypeText("\t")
#
#   myRange[["Start"]] <- bm[["Range"]][["Start"]]
#   myRange$Select()
#   bm$Delete()
#   wrd[["Selection"]]$ConvertToTable(Separator       = wdConst$wdSeparateByTabs,
#                                     NumColumns      = dim1,
#                                     NumRows         = dim2,
#                                     AutoFitBehavior = wdConst$wdAutoFitFixed)
#
#   wrdTable <- wrd[["Selection"]][["Tables"]]$Item(1)
#   # http://www.thedoctools.com/downloads/DocTools_List_Of_Built-in_Style_English_Danish_German_French.pdf
#   wrdTable[["Style"]] <- -115 # "Tabelle Klassisch 1"
#   wrdSel <- wrd[["Selection"]]
#
#
#   # align the columns
#   if(is.null(align))
#     align <- c("l", rep(x = "r", ncol(tab)-1))
#   else
#     align <- rep(align, length.out=ncol(tab))
#
#   align[align=="l"] <- wdConst$wdAlignParagraphLeft
#   align[align=="c"] <- wdConst$wdAlignParagraphCenter
#   align[align=="r"] <- wdConst$wdAlignParagraphRight
#
#   for(i in seq_along(align)){
#     wrdTable$Columns(i)$Select()
#     wrd[["Selection"]][["ParagraphFormat"]][["Alignment"]] <- align[i]
#   }
#
#   if(!is.null(font)){
#     wrdTable$Select()
#     WrdFont(wrd) <- font
#   }
#
#   if(autofit)
#     wrdTable$Columns()$AutoFit()
#
#   # Cursor aus der Tabelle auf die letzte Postition im Dokument setzten
#   # Selection.GoTo What:=wdGoToPercent, Which:=wdGoToLast
#   wrd[["Selection"]]$GoTo(What = wdConst$wdGoToPercent, Which= wdConst$wdGoToLast)
#
#   if(!is.null(main)){
#     # insert caption
#     sel <- wrd$Selection()  # "Abbildung"
#     sel$InsertCaption(Label=wdConst$wdCaptionTable, Title=main)
#     sel$TypeParagraph()
#
#   }
#
#   invisible(wrdTable)
#
# }
#

# WrdTable <- function(tab, wrd){

# ###  http://home.wanadoo.nl/john.hendrickx/statres/other/PasteAsTable.html

# write.table(tab, file="clipboard", sep="\t", quote=FALSE)

# myRange <- wrd[["Selection"]][["Range"]]

# bm <- wrd[["ActiveDocument"]][["Bookmarks"]]$Add("PasteHere", myRange)

# myRange$Paste()
# wrd[["Selection"]]$TypeText("\t")

# myRange[["Start"]] <- bm[["Range"]][["Start"]]
# myRange$Select()

# bm$Delete()

# wrd[["Selection"]]$ConvertToTable(Separator=wdConst$wdSeparateByTabs, NumColumns=4,
# NumRows=9, AutoFitBehavior=wdConst$wdAutoFitFixed)

# wrdTable <- wrd[["Selection"]][["Tables"]]$Item(1)
# wrdTable[["Style"]] <- "Tabelle Klassisch 1"

# wrdSel <- wrd[["Selection"]]
# wrdSel[["ParagraphFormat"]][["Alignment"]] <- wdConst$wdAlignParagraphRight

# #left align the first column
# wrdTable[["Columns"]]$Item(1)$Select()
# wrd[["Selection"]][["ParagraphFormat"]][["Alignment"]] <- wdConst$wdAlignParagraphLeft

# ### wtab[["ApplyStyleHeadingRows"]] <- TRUE
# ### wtab[["ApplyStyleLastRow"]] <- FALSE
# ### wtab[["ApplyStyleFirstColumn"]] <- TRUE
# ### wtab[["ApplyStyleLastColumn"]] <- FALSE
# ### wtab[["ApplyStyleRowBands"]] <- TRUE
# ### wtab[["ApplyStyleColumnBands"]] <- FALSE

# ### With Selection.Tables(1)
# #### If .Style <> "Tabellenraster" Then
# ### .Style = "Tabellenraster"
# ### End If

# ### wrd[["Selection"]]$ConvertToTable( Separator=wdConst$wdSeparateByTabs, AutoFit=TRUE, Format=wdConst$wdTableFormatSimple1,
# ### ApplyBorders=TRUE, ApplyShading=TRUE, ApplyFont=TRUE,
# ### ApplyColor=TRUE, ApplyHeadingRows=TRUE, ApplyLastRow=FALSE,
# ### ApplyFirstColumn=TRUE, ApplyLastColumn=FALSE)

# ###  wrd[["Selection"]][["Tables"]]$Item(1)$Select()
# #wrd[["Selection"]][["ParagraphFormat"]][["Alignment"]] <- wdConst$wdAlignParagraphRight
# ### ### left align the first column
# ### wrd[["Selection"]][["Columns"]]$Item(1)$Select()
# ### wrd[["Selection"]][["ParagraphFormat"]][["Alignment"]] <- wdConst$wdAlignParagraphLeft
# ### wrd[["Selection"]][["ParagraphFormat"]][["Alignment"]] <- wdConst$wdAlignParagraphRight



# }




# require ( xtable )
# data ( tli )
# fm1 <- aov ( tlimth ~ sex + ethnicty + grade + disadvg , data = tli )
# fm1.table <- print ( xtable (fm1), type ="html")

# Tabellen-Studie via HTML FileExport


# WrdInsTable <- function( tab, wrd ){
# htmtab <- print(xtable(tab), type ="html")

# ### Let's create a summary file and insert it
# ### get a tempfile:
# fn <- paste(tempfile(pattern = "file", tmpdir = tempdir()), ".txt", sep="")

# write(htmtab, file=fn)
# wrd[["Selection"]]$InsertFile(fn)
# wrd[["ActiveDocument"]][["Tables"]]$Item(
# wrd[["ActiveDocument"]][["Tables"]][["Count"]] )[["Style"]] <- "Tabelle Klassisch 1"

# }

# WrdInsTable( fm1, wrd=wrd )

# data(d.pizza)
# txt <- Desc( temperature ~ driver, data=d.pizza )
# WrdInsTable( txt, wrd=wrd )

# WrdPlot(PlotDescNumFact( temperature ~ driver, data=d.pizza, newwin=TRUE )
# , wrd=wrd, width=17, crop=c(0,0,60,0))




## Entwicklungs-Ideen ====


# With ActiveDocument.Bookmarks
# .Add Range:=Selection.Range, Name:="start"
# .DefaultSorting = wdSortByName
# .ShowHidden = False
# End With
# Selection.TypeText Text:="Hier kommt mein Text"
# Selection.TypeParagraph
# Selection.TypeText Text:="und auf weiteren Zeilen"
# Selection.TypeParagraph
# With ActiveDocument.Bookmarks
# .Add Range:=Selection.Range, Name:="stop"
# .DefaultSorting = wdSortByName
# .ShowHidden = False
# End With
# Selection.GoTo What:=wdGoToBookmark, Name:="start"
# Selection.GoTo What:=wdGoToBookmark, Name:="stop"
# With ActiveDocument.Bookmarks
# .DefaultSorting = wdSortByName
# .ShowHidden = False
# End With
# Selection.MoveLeft Unit:=wdWord, Count:=2, Extend:=wdExtend
# Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
# Selection.Font.Name = "Arial Black"
# Selection.EndKey Unit:=wdStory
# Selection.GoTo What:=wdGoToBookmark, Name:="stop"
# Selection.Find.ClearFormatting
# With Selection.Find
# .Text = "0."
# .Replacement.Text = " ."
# .Forward = True
# .Wrap = wdFindContinue
# .Format = False
# .MatchCase = False
# .MatchWholeWord = False
# .MatchWildcards = False
# .MatchSoundsLike = False
# .MatchAllWordForms = False
# End With
# ActiveDocument.Bookmarks("start").Delete
# With ActiveDocument.Bookmarks
# .DefaultSorting = wdSortByName
# .ShowHidden = False
# End With
# End Sub
# wdSortByName =0
# wdGoToBookmark = -1
# wdFindContinue = 1
# wdStory = 6



# Bivariate Darstellungen gute uebersicht
# pairs( lapply( lapply( c( d.set[,-1], list()), "as.numeric" ), "jitter" ), col=rgb(0,0,0,0.2) )


# Gruppenweise Mittelwerte fuer den ganzen Recordset
# wrdInsertText( "Mittelwerte zusammengefasst\n\n" )
# wrdInsertSummary(
# signif( cbind(
# t(as.data.frame( lapply( d.frm, tapply, grp, "mean", na.rm=TRUE )))
# , tot=mean(d.frm, na.rm=TRUE)
# ), 3)

