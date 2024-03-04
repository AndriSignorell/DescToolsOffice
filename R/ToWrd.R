

# ToWrd


#' Send Objects to Word
#' 
#' Send objects like tables, ftables, lm tables, TOnes or just simple texts to
#' a MS-Word document.
#' 
#' The paragraph format can be defined by means of these properties:
#' 
#' \code{LeftIndent}, \code{RightIndent}, \code{SpaceBefore},
#' \code{SpaceBeforeAuto}, \code{SpaceAfter}, \code{SpaceAfterAuto},
#' \code{LineSpacingRule}, \code{Alignment}, \code{WidowControl},
#' \code{KeepWithNext}, \code{KeepTogether}, \code{PageBreakBefore},
#' \code{NoLineNumber}, \code{Hyphenation}, \code{FirstLineIndent},
#' \code{OutlineLevel}, \code{CharacterUnitLeftIndent},
#' \code{CharacterUnitRightIndent}, \code{CharacterUnitFirstLineIndent},
#' \code{LineUnitBefore}, \code{LineUnitAfter}, \code{MirrorIndents}.
#' 
#' @aliases ToWrd ToWrd.table ToWrd.ftable ToWrd.character ToWrd.lm ToWrd.TOne
#' ToWrd.TMod ToWrd.Freq ToWrd.default ToWrd.data.frame
#' @param x the object to be transferred to Word.
#' @param font the font to be used to the output. This should be defined as a
#' list containing fontname, fontsize, bold and italic flags:\cr
#' \code{list(name="Arial", size=10, bold=FALSE, italic=TRUE)}.
#' @param para list containing paragraph format properties to be applied to the
#' inserted text. For right align the paragraph one can set: \cr
#' \code{list(alignment="r", LineBefore=0.5)}. See details for the full set of
#' properties.
#' @param main a caption for a table. This will be inserted by
#' \code{\link{WrdCaption}} in Word and can be listed afterwards in a specific
#' index. Default is \code{NULL}, which will insert nothing. Ignored if
#' \code{x} is not a table.
#' @param align character vector giving the alignment of the table columns.
#' \code{"l"} means left, \code{"r"} right and \code{"c"} center alignement.
#' The code will be recyled to the length of thenumber of columns.
#' @param method string specifying how the \code{"ftable"} object is formatted
#' (and printed if used as in \code{write.ftable()} or the \code{print}
#' method).  Can be abbreviated.  Available methods are (see the examples):
#' \describe{ \item{list("\"non.compact\"")}{the default representation of an
#' \code{"ftable"} object.} \item{list("\"row.compact\"")}{a row-compact
#' version without empty cells below the column labels.}
#' \item{list("\"col.compact\"")}{a column-compact version without empty cells
#' to the right of the row labels.} \item{list("\"compact\"")}{a row- and
#' column-compact version.  This may imply a row and a column label sharing the
#' same cell.  They are then separated by the string \code{lsep}.} }
#' @param autofit logical, defining if the columns of table should be fitted to
#' the length of their content.
#' @param row.names logical, defining whether the row.names should be included
#' in the output. Default is \code{FALSE}.
#' @param col.names logical, defining whether the col.names should be included
#' in the output. Default is \code{TRUE}.
#' @param tablestyle either the name of a defined Word tablestyle or its index.
#' @param style character, name of a style to be applied to the inserted text.
#' @param \dots further arguments to be passed to or from methods.
#' @param bullet logical, defines if the text should be formatted as bullet
#' points.
#' @param split character vector (or object which can be coerced to such)
#' containing regular expression(s) (unless \code{fixed = TRUE}) to use for
#' splitting. If empty matches occur, in particular if \code{split} has length
#' 0, x is split into single characters. If \code{split} has length greater
#' than 1, it is re-cycled along x.
#' @param fixed logical. If TRUE match split exactly, otherwise use regular
#' expressions. Has priority over perl.
#' @param digits integer, the desired (fixed) number of digits after the
#' decimal point. Unlike \code{\link{formatC}} you will always get this number
#' of digits even if the last digit is 0.
#' @param na.form character, string specifying how \code{NA}s should be
#' specially formatted.  If set to \code{NULL} (default) no special action will
#' be taken.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return if \code{x} is a table a pointer to the table will be returned
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{GetNewWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' # we can't get this through the CRAN test - run it with copy/paste to console
#' 
#' wrd <- GetNewWrd()
#' ToWrd("This is centered Text in Arial Black\n",
#'       para=list(Alignment=wdConst$wdAlignParagraphCenter,
#'                 SpaceBefore=3, SpaceAfter=6),
#'       font=list(name="Arial Black", size=14),
#'       wrd=wrd)
#' 
#' sel <- wrd$Selection()$Borders(wdConst$wdBorderBottom)
#' sel[["LineStyle"]] <- wdConst$wdLineStyleSingle
#' 
#' 
#' t1 <- TOne(x = d.pizza[, c("temperature","delivery_min","driver","wine_ordered")],
#'            grp=d.pizza$wine_delivered)
#' 
#' ToWrd(t1, font=list(name="Algerian"), wrd=wrd)
#' 
#' 
#' tab <- table(d.pizza$driver, d.pizza$area)
#' 
#' tab <- table(d.pizza$driver, d.pizza$area)
#' ToWrd(tab, font = list(size=15, name="Arial"), row.names = TRUE, col.names = TRUE,
#'       main= "my Title", wrd=wrd)
#' ToWrd(tab, font = list(size=10, name="Arial narrow"),
#'       row.names = TRUE, col.names=FALSE, wrd=wrd)
#' ToWrd(tab, font = list(size=15, name="Arial"), align="r",
#'       row.names = FALSE, col.names=TRUE, wrd=wrd)
#' ToWrd(tab, font = list(size=15, name="Arial"),
#'       row.names = FALSE, col.names=FALSE, wrd=wrd)
#' 
#' ToWrd(tab, tablestyle = "Mittlere Schattierung 2 - Akzent 4",
#'       row.names=TRUE, col.names=TRUE, wrd=wrd)
#' 
#' ToWrd(Format(tab, big.mark = "'", digits=0), wrd=wrd)
#' 
#' zz <- ToWrd(Format(tab, big.mark = "'", digits=0), wrd=wrd)
#' zz$Rows(1)$Select()
#' WrdFont(wrd = wrd) <- list(name="Algerian", size=14, bold=TRUE)
#' 
#' 
#' # Send a TMod table to Word using a split to separate columns
#' r.ols <- lm(Fertility ~ . , swiss)
#' r.gam <- glm(Fertility ~ . , swiss, family=Gamma(link="identity"))
#' 
#' # Build the model table for some two models, creating a user defined
#' # reporting function (FUN) with | as column splitter
#' tm <- TMod(OLS=r.ols, Gamma=r.gam, 
#'            FUN=function(est, se, tval, pval, lci, uci){
#'               gettextf("%s|[%s, %s]|%s",
#'                        Format(est, fmt=Fmt("num"), digits=2),
#'                        Format(lci, fmt=Fmt("num"), digits=2), 
#'                        Format(uci, fmt=Fmt("num"), digits=2),
#'                        Format(pval, fmt="*")
#'               )})
#' 
#' # send it to Word, where we get a table with 3 columns per model
#' # coef | confint | p-val
#' wrd <- GetNewWrd()
#' ToWrd(tm, split="|", align=StrSplit("lrclrcl"))
#' )
#' }
#' 
#' @export ToWrd
ToWrd <- function(x, font=NULL, ..., wrd=DescToolsOfficeOptions("lastWord")){
  UseMethod("ToWrd")
}


# ToWrdB <- function(x, font = NULL, ..., wrd = DescToolsOfficeOptions("lastWord"), 
#                     bookmark=gettextf("b%s", sample(1e9, 1))){
#   
#   bm <- WrdInsertBookmark(name = bookmark, wrd=wrd)
#   ToWrd(x, font=font, ..., wrd=wrd)
#   
#   d <- wrd$Selection()$range()$start() - bm$range()$start()
#   wrd$Selection()$MoveLeft(Unit=wdConst$wdCharacter, Count=d, Extend=wdConst$wdExtend)
#   
#   bm <- WrdInsertBookmark(name = bookmark, wrd=wrd)
#   
#   wrd[["Selection"]]$Collapse(Direction=wdConst$wdCollapseEnd)
#   
#   invisible(bm)
#   
# }


# function to generate random bookmark names 
# (ensure we'll always get 9 digits with min=0.1)
.randbm <- function() paste("bm", round(runif(1, min=0.1)*1e9), sep="")





#' Send Objects to Word and Bookmark Them
#' 
#' Send objects like tables, ftables, lm tables, TOnes or just simple texts to
#' a MS-Word document and place a bookmark on them. This has the advantage,
#' that objects in a Word document can be updated later, provided the bookmark
#' name has been stored.
#' 
#' This function encapsulates \code{\link{ToWrd}}, by placing a bookmark over
#' the complete inserted results. The given name can be questioned with
#' \code{bm$name()}.
#' 
#' @param x the object to be transferred to Word.
#' @param font the font to be used to the output. This should be defined as a
#' list containing fontname, fontsize, bold and italic flags:\cr
#' \code{list(name="Arial", size=10, bold=FALSE, italic=TRUE)}.
#' @param \dots further arguments to be passed to or from methods.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @param bookmark the name of the bookmark.
#' @return a handle to the set bookmark
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{ToWrd}}, \code{\link{WrdInsertBookmark}}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' # we can't get this through the CRAN test - run it with copy/paste to console
#' 
#' wrd <- GetNewWrd()
#' bm <- ToWrdB("This is text to be possibly replaced later.")
#' 
#' # get the automatically created name of the bookmark
#' bm$name()
#' 
#' WrdGoto(bm$name())
#' UpdateBookmark(...)
#' }
#' 
#' @export ToWrdB
ToWrdB <- function(x, font = NULL, ..., wrd = DescToolsOfficeOptions("lastWord"), 
                   bookmark=gettextf("bmt%s", round(runif(1, min=0.1)*1e9))){
  
  # Sends the output of an object x to word and places a bookmark bm on it
  
  # place the temporary bookmark on cursor
  bm_start <- WrdInsertBookmark(.randbm())
  
  # send stuff to Word (it's generic ...)
  ToWrd(x, font=font, ..., wrd=wrd)
  
  # place end bookmark
  bm_end <- WrdInsertBookmark(.randbm())
  
  # select all the inserted text between the two bookmarks
  wrd[["ActiveDocument"]]$Range(bm_start$range()$start(), bm_end$range()$end())$select()
  
  # place the required bookmark over the whole inserted story
  res <- WrdInsertBookmark(bookmark)
  
  # collapse selection to the end position
  wrd$selection()$collapse(wdConst$wdCollapseEnd)
  
  # delete the two temporary bookmarks start/end
  bm_start$delete()
  bm_end$delete()
  
  # return the bookmark with inserted story
  invisible(res)
  
}




#' Send a Plot to Word and Bookmark it
#' 
#' Evaluate given plot code to a \code{\link{tiff}()} device and imports the
#' created plot in the currently open MS-Word document. The imported plot is
#' marked with a bookmark that can later be used for a potential update
#' (provided the bookmark name has been stored).
#' 
#' An old and persistent problem that has existed for a long time is that as
#' results once were loaded into a Word document the connection broke so that
#' no update was possible. It was only recently that I realized that bookmarks
#' in Word could be a solution for this. The present function evaluates some
#' given plot code chunk using a tiff device and imports the created plot in a
#' word document. The imported plot is given a bookmark, that can be used
#' afterwards for changing or updating the plot.
#' 
#' This function is designed for use with the \bold{DescToolsAddIns} functions
#' \code{ToWrdPlotWithBookmark()} and \code{ToWrdWithBookmark()} allowing to
#' assign keyboard shortcuts. The two functions will also insert the newly
#' defined bookmark in the source file in a format, which can be interpreted by
#' the function \code{UpdateBookmark()}.
#' 
#' @param plotcode code chunk needed for producing the plot
#' @param bookmark character, the name of the bookmark
#' @param width the width in cm of the plot in the Word document (default 15)
#' @param height the height in cm of the plot in the Word document (default
#' 9.3)
#' @param scale the scale of the plot (default 100)
#' @param pointsize the default pointsize of plotted text, interpreted as big
#' points (1/72 inch) at \code{res} ppi. (default is 12)
#' @param res the resolution for the graphic (default 300)
#' @param crop a vector of 4 elements, the crop factor for all 4 sides of a
#' picture in cm (default all 0)
#' @param title character, the title of the plot to be inserted in the word
#' document
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOfficeOptions("lastWord")}.
#' @return a list \item{plot_hwnd }{a windows handle to the inserted plot}
#' \item{bookmark}{a windows handle to the bookmark}
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{ToWrdB}}, \code{\link{WrdInsertBookmark}}
#' @keywords print
#' @examples
#' 
#' \dontrun{
#' # we can't get this through the CRAN test - run it with copy/paste to console
#' 
#' wrd <- GetNewWrd()
#' bm <- ToWrdB("This is text to be possibly replaced later.")
#' 
#' # get the automatically created name of the bookmark
#' bm$name()
#' 
#' WrdGoto(bm$name())
#' UpdateBookmark(...)
#' }
#' 
#' @export ToWrdPlot
ToWrdPlot <- function(plotcode,  
                      width=NULL, height=NULL, scale=100, pointsize=12, res=300, crop=0, title=NULL, 
                      wrd = DescToolsOfficeOptions("lastWord"), 
                      bookmark=gettextf("bmp%s", round(runif(1, min=0.1)*1e9))
){
  
  if(is.null(width)) width <- 15
  if(is.null(height)) height <- width / gold_sec_c 
  
  crop <- rep(crop, length.out=4)
  
  if(is.null(bookmark)) bookmark <- .randbm()
  
  
  # open device
  tiff(filename = (fn <- paste(tempfile(), ".tif", sep = "")), 
       width = width, height = height, units = "cm", pointsize = pointsize,
       res = res, compression = "lzw")
  
  # do plot
  if(!is.null(plotcode ))
    eval(parse(text = plotcode))
  
  # close device
  dev.off()
  
  
  # import in word ***********
  # place the temporary bookmark on cursor
  bm_start <- WrdInsertBookmark(.randbm(), wrd=wrd)
  
  # send stuff to Word (it's generic ...)
  hwnd <- wrd$selection()$InlineShapes()$AddPicture(FileName=fn, LinkToFile=FALSE, SaveWithDocument=TRUE)
  hwnd[["LockAspectRatio"]] <- 1
  hwnd[["ScaleWidth"]] <- hwnd[["ScaleHeight"]] <- scale
  pic <- hwnd$PictureFormat()
  pic[["CropBottom"]] <- CmToPts(crop[1])
  pic[["CropLeft"]] <- CmToPts(crop[2])
  pic[["CropTop"]] <- CmToPts(crop[3])
  pic[["CropRight"]] <- CmToPts(crop[4])
  
  if(!is.null(title)){
    hwnd$select()
    wrd[["Selection"]]$InsertCaption(Label="Figure", Title=gettextf(" - %s", title), 
                                     Position=wdConst$wdCaptionPositionBelow, ExcludeLabel=0)
    wrd[["Selection"]]$MoveRight(wdConst$wdCharacter, 1, 0)
    
  }
  
  
  ToWrd(x="\n", wrd=wrd)
  
  # place end bookmark
  bm_end <- WrdInsertBookmark(.randbm(), wrd=wrd)
  
  # select all the inserted text between the two bookmarks
  wrd[["ActiveDocument"]]$Range(bm_start$range()$start(), bm_end$range()$end())$select()
  
  # place the required bookmark over the whole inserted story
  res <- WrdInsertBookmark(bookmark, wrd=wrd)
  
  # collapse selection to the end position
  wrd$selection()$collapse(wdConst$wdCollapseEnd)
  
  # delete the two temporary bookmarks start/end
  bm_start$delete()
  bm_end$delete()
  
  # return the bookmark with inserted story
  invisible(list(plot_hwnd=hwnd, bookmark=res))
  
}






#' @rdname ToWrd
ToWrd.default <- function(x, font=NULL, ..., wrd=DescToolsOfficeOptions("lastWord")){
  
  ToWrd.character(x=DescTools:::.CaptOut(x), font=font, ..., wrd=wrd)
  invisible()
  
}



ToWrd.Desc <- function(x, font=NULL, ..., wrd=DescToolsOfficeOptions("lastWord")){
  
  printWrd(x, ..., wrd=wrd)
  invisible()
  
}



#' @rdname ToWrd
ToWrd.TOne <- function(x, font=NULL, para=NULL, main=NULL, align=NULL,
                       autofit=TRUE, ..., wrd=DescToolsOfficeOptions("lastWord")){
  
  wTab <- ToWrd.table(x, main=NULL, font=font, align=align, autofit=autofit, wrd=wrd, ...)
  
  if(!is.null(para)){
    wTab$Select()
    WrdParagraphFormat(wrd) <- para
    
    # move out of table
    wrd[["Selection"]]$EndOf(wdConst$wdTable)
    wrd[["Selection"]]$MoveRight(wdConst$wdCharacter, 2, 0)
  }
  
  if(is.null(font)) font <- list()
  if(is.null(font$size))
    font$size <- WrdFont(wrd)$size - 2
  else
    font$size <- font$size - 2
  
  wrd[["Selection"]]$TypeBackspace()
  ToWrd.character(paste("\n", attr(x, "legend"), "\n\n", sep=""),
                  font=font, wrd=wrd)
  
  
  if(!is.null(main)){
    sel <- wrd$Selection()  # "Abbildung"
    sel$InsertCaption(Label=wdConst$wdCaptionTable, Title=paste(" - ", main, sep=""))
    sel$TypeParagraph()
    
  }
  
  invisible(wTab)
  
}


#' @rdname ToWrd
ToWrd.abstract <- function(x, font=NULL, autofit=TRUE, ..., wrd=DescToolsOfficeOptions("lastWord")){
  
  WrdCaption(x=attr(x, "main"), wrd=wrd)
  
  if(!is.null(attr(x, "label"))){
    
    if(is.null(font)){
      lblfont <- list(fontsize=8)
    } else {
      lblfont <- font
      lblfont$fontsize <- 8
    }
    
    ToWrd.character(paste("\n", attr(x, "label"), "\n", sep=""),
                    font = lblfont, wrd=wrd)
  }
  
  ToWrd.character(gettextf("\ndata.frame:	%s obs. of  %s variables (complete cases: %s / %s)\n\n",
                           attr(x, "nrow"), attr(x, "ncol"), attr(x, "complete"), 
                           DescTools::Format(attr(x, "complete")/attr(x, "nrow"), fmt="%", digits=1))
                  , font=font, wrd=wrd)
  
  wTab <- ToWrd.data.frame(x, wrd=wrd, autofit=autofit, font=font, align="l", ...)
  
  invisible(wTab)
  
}



ToWrd.lm <- function(x, font=NULL, ..., wrd=DescToolsOfficeOptions("lastWord")){
  
  invisible()
}



#' @rdname ToWrd
ToWrd.character <- function (x, font = NULL, para = NULL, style = NULL, bullet=FALSE,  ..., wrd = DescToolsOfficeOptions("lastWord")) {
  
  # we will convert UTF-8 strings to Latin-1, if the local info is Latin-1
  if (any(l10n_info()[["Latin-1"]] & Encoding(x) == "UTF-8"))
    x[Encoding(x) == "UTF-8"] <- iconv(x[Encoding(x) == "UTF-8"], from = "UTF-8", to = "latin1")
  
  wrd[["Selection"]]$InsertAfter(paste(x, collapse = "\n"))
  
  if (!is.null(style))
    WrdStyle(wrd) <- style
  
  if (!is.null(para))
    WrdParagraphFormat(wrd) <- para
  
  
  if(identical(font, "fix")){
    font <- DescToolsOfficeOptions("fixedfont")
    if(is.null(font))
      font <- structure(list(name="Courier New", size=8), class="font")
  }
  
  if(!is.null(font)){
    currfont <- WrdFont(wrd)
    WrdFont(wrd) <- font
    on.exit(WrdFont(wrd) <- currfont)
  }
  
  if(bullet)
    wrd[["Selection"]]$Range()$ListFormat()$ApplyBulletDefault()
  
  wrd[["Selection"]]$Collapse(Direction=wdConst$wdCollapseEnd)
  
  invisible()
  
}


ToWrd.PercTable <- function(x, font=NULL, main = NULL, ..., wrd = DescToolsOfficeOptions("lastWord")){
  ToWrd.ftable(x$ftab, font=font, main=main, ..., wrd=wrd)
}



ToWrd.data.frame <- function(x, font=NULL, main = NULL, row.names=NULL, ..., wrd = DescToolsOfficeOptions("lastWord")){
  
  # drops dimension names!! don't use here
  # x <- apply(x, 2, as.character)
  
  x[] <- lapply(x, as.character)
  x <- as.matrix(x)
  
  if(is.null(row.names))
    if(identical(row.names(x), as.character(1:nrow(x))))
      row.names <- FALSE
  else
    row.names <- TRUE
  
  ToWrd.table(x=x, font=font, main=main, row.names=row.names, ..., wrd=wrd)
}


# ToWrd.data.frame <- function(x, font=NULL, main = NULL, row.names=NULL, as.is=FALSE, ..., wrd = DescToolsOfficeOptions("lastWord")){
#
#   if(as.is)
#     x <- apply(x, 2, as.character)
#   else
#     x <- FixToTable(capture.output(x))
#
#   if(is.null(row.names))
#     if(identical(row.names, seq_along(1:nrow(x))))
#       row.names <- FALSE
#     else
#       row.names <- TRUE
#
#     if(row.names==TRUE)
#       x <- cbind(row.names(x), x)
#
#     ToWrd.table(x=x, font=font, main=main, ..., wrd=wrd)
# }


ToWrd.matrix <- function(x, font=NULL, main = NULL, ..., wrd = DescToolsOfficeOptions("lastWord")){
  ToWrd.table(x=x, font=font, main=main, ..., wrd=wrd)
}


ToWrd.Freq <- function(x, font=NULL, main = NULL, ..., wrd = DescToolsOfficeOptions("lastWord")){
  
  x[,c(3,5)] <- sapply(round(x[,c(3,5)], 3), Format, digits=3)
  
  res <- ToWrd.data.frame(x=x, main=main, font=font, wrd=wrd)
  
  invisible(res)
  
}




ToWrd.ftable <- function (x, font = NULL, main = NULL, align=NULL, method = "compact", ..., wrd = DescToolsOfficeOptions("lastWord")) {
  
  # simple version:
  #   x <- FixToTable(capture.output(x))
  #   ToWrd.character(x, font=font, main=main, ..., wrd=wrd)
  
  # let R do all the complicated formatting stuff
  # but we can't import a not exported function, so we provide an own copy of it
  
  # so this is a verbatim copy of it
  .format.ftable <- function (x, quote = TRUE, digits = getOption("digits"), method = c("non.compact",
                                                                                        "row.compact", "col.compact", "compact"), lsep = " | ", ...)
  {
    if (!inherits(x, "ftable"))
      stop("'x' must be an \"ftable\" object")
    charQuote <- function(s) if (quote && length(s))
      paste0("\"", s, "\"")
    else s
    makeLabels <- function(lst) {
      lens <- lengths(lst)
      cplensU <- c(1, cumprod(lens))
      cplensD <- rev(c(1, cumprod(rev(lens))))
      y <- NULL
      for (i in rev(seq_along(lst))) {
        ind <- 1 + seq.int(from = 0, to = lens[i] - 1) *
          cplensD[i + 1L]
        tmp <- character(length = cplensD[i])
        tmp[ind] <- charQuote(lst[[i]])
        y <- cbind(rep(tmp, times = cplensU[i]), y)
      }
      y
    }
    makeNames <- function(x) {
      nmx <- names(x)
      if (is.null(nmx))
        rep_len("", length(x))
      else nmx
    }
    l.xrv <- length(xrv <- attr(x, "row.vars"))
    l.xcv <- length(xcv <- attr(x, "col.vars"))
    method <- match.arg(method)
    if (l.xrv == 0) {
      if (method == "col.compact")
        method <- "non.compact"
      else if (method == "compact")
        method <- "row.compact"
    }
    if (l.xcv == 0) {
      if (method == "row.compact")
        method <- "non.compact"
      else if (method == "compact")
        method <- "col.compact"
    }
    LABS <- switch(method, non.compact = {
      cbind(rbind(matrix("", nrow = length(xcv), ncol = length(xrv)),
                  charQuote(makeNames(xrv)), makeLabels(xrv)), c(charQuote(makeNames(xcv)),
                                                                 rep("", times = nrow(x) + 1)))
    }, row.compact = {
      cbind(rbind(matrix("", nrow = length(xcv) - 1, ncol = length(xrv)),
                  charQuote(makeNames(xrv)), makeLabels(xrv)), c(charQuote(makeNames(xcv)),
                                                                 rep("", times = nrow(x))))
    }, col.compact = {
      cbind(rbind(cbind(matrix("", nrow = length(xcv), ncol = length(xrv) -
                                 1), charQuote(makeNames(xcv))), charQuote(makeNames(xrv)),
                  makeLabels(xrv)))
    }, compact = {
      xrv.nms <- makeNames(xrv)
      xcv.nms <- makeNames(xcv)
      mat <- cbind(rbind(cbind(matrix("", nrow = l.xcv - 1,
                                      ncol = l.xrv - 1), charQuote(makeNames(xcv[-l.xcv]))),
                         charQuote(xrv.nms), makeLabels(xrv)))
      mat[l.xcv, l.xrv] <- paste(tail(xrv.nms, 1), tail(xcv.nms,
                                                        1), sep = lsep)
      mat
    }, stop("wrong method"))
    DATA <- rbind(if (length(xcv))
      t(makeLabels(xcv)), if (method %in% c("non.compact",
                                            "col.compact"))
        rep("", times = ncol(x)), format(unclass(x), digits = digits,
                                         ...))
    cbind(apply(LABS, 2L, format, justify = "left"), apply(DATA,
                                                           2L, format, justify = "right"))
  }
  
  
  tab <- .format.ftable(x, quote=FALSE, method=method, lsep="")
  tab <- StrTrim(tab)
  
  if(is.null(align))
    align <- c(rep("l", length(attr(x, "row.vars"))), rep("r", ncol(x)))
  
  wtab <- ToWrd.table(tab, font=font, main=main, align=align, ..., wrd=wrd)
  
  invisible(wtab)
  
}



#' @rdname ToWrd
ToWrd.table <- function (x, font = NULL, main = NULL, align=NULL, tablestyle=NULL, autofit = TRUE,
                         row.names=TRUE, col.names=TRUE, ..., wrd = DescToolsOfficeOptions("lastWord")) {
  
  
  x[] <- as.character(x)
  if (any(l10n_info()[["Latin-1"]] & Encoding(x) == "UTF-8"))
    x[Encoding(x) == "UTF-8"] <- iconv(x[Encoding(x) == "UTF-8"], from = "UTF-8", to = "latin1")
  
  # add column names to character table
  if(col.names)
    x <- rbind(colnames(x), x)
  if(row.names){
    rown <- rownames(x)
    # if(col.names)
    #   rown <- c("", rown)
    x <- cbind(rown, x)
  }
  # replace potential \n in table with /cr, as convertToTable would make a new cell for them
  x <- gsub(pattern= "\n", replacement = "/cr", x = x)
  # paste the cells and separate by \t
  txt <- paste(apply(x, 1, paste, collapse="\t"), collapse="\n")
  
  nc <- ncol(x)
  nr <- nrow(x)
  
  # insert and convert
  wrd[["Selection"]]$InsertAfter(txt)
  wrdTable <- wrd[["Selection"]]$ConvertToTable(Separator = wdConst$wdSeparateByTabs,
                                                NumColumns = nc,  NumRows = nr,
                                                AutoFitBehavior = wdConst$wdAutoFitFixed)
  
  wrdTable[["ApplyStyleHeadingRows"]] <- col.names
  
  # replace /cr by \n again in word
  wrd[["Selection"]][["Find"]]$ClearFormatting()
  wsel <- wrd[["Selection"]][["Find"]]
  wsel[["Text"]] <- "/cr"
  wrep <- wsel[["Replacement"]]
  wrep[["Text"]] <- "^l"
  wsel$Execute(Replace=wdConst$wdReplaceAll)
  
  
  # http://www.thedoctools.com/downloads/DocTools_List_Of_Built-in_Style_English_Danish_German_French.pdf
  if(is.null(tablestyle)){
    WrdTableBorders(wrdTable, from=c(1,1), to=c(1, nc),
                    border = wdConst$wdBorderTop)
    if(col.names)
      WrdTableBorders(wrdTable, from=c(1,1), to=c(1, nc),
                      border = wdConst$wdBorderBottom)
    
    WrdTableBorders(wrdTable, from=c(nr, 1), to=c(nr, nc),
                    border = wdConst$wdBorderBottom)
    
    space <- RoundTo((if(is.null(font$size)) WrdFont(wrd)$size else font$size) * .2, multiple = .5)
    wrdTable$Rows(1)$Select()
    WrdParagraphFormat(wrd) <- list(SpaceBefore=space, SpaceAfter=space)
    
    if(col.names){
      wrdTable$Rows(2)$Select()
      WrdParagraphFormat(wrd) <- list(SpaceBefore=space)
    }
    
    wrdTable$Rows(nr)$Select()
    WrdParagraphFormat(wrd) <- list(SpaceAfter=space)
    
    # wrdTable[["Style"]] <- -115 # code for "Tabelle Klassisch 1"
  } else
    if(!is.na(tablestyle))
      wrdTable[["Style"]] <- tablestyle
  
  
  # align the columns
  if(is.null(align))
    align <- c(rep("l", row.names), rep(x = "r", nc-row.names))
  else
    align <- rep(align, length.out=nc)
  
  align[align=="l"] <- wdConst$wdAlignParagraphLeft
  align[align=="c"] <- wdConst$wdAlignParagraphCenter
  align[align=="r"] <- wdConst$wdAlignParagraphRight
  
  for(i in seq_along(align)){
    wrdTable$Columns(i)$Select()
    wrdSel <- wrd[["Selection"]]
    wrdSel[["ParagraphFormat"]][["Alignment"]] <- align[i]
  }
  
  if(!is.null(font)){
    wrdTable$Select()
    WrdFont(wrd) <- font
  }
  
  if(autofit)
    wrdTable$Columns()$AutoFit()
  
  
  # this will get us out of the table and put the text cursor directly behind it
  wrdTable$Select()
  wrd[["Selection"]]$Collapse(wdConst$wdCollapseEnd)
  
  # instead of coarsely moving to the end of the document ...
  # Selection.GoTo What:=wdGoToPercent, Which:=wdGoToLast
  # wrd[["Selection"]]$GoTo(What = wdConst$wdGoToPercent, Which= wdConst$wdGoToLast)
  
  if(!is.null(main)){
    # insert caption
    sel <- wrd$Selection()  
    sel$InsertCaption(Label=wdConst$wdCaptionTable, Title=paste(" - ", main, sep=""))
    sel$TypeParagraph()
    
  }
  
  wrd[["Selection"]]$TypeParagraph()
  
  invisible(wrdTable)
  
}





ToWrd.TwoGroups <- function(x, font = NULL, ..., 
                            wrd = DescToolsOfficeOptions("lastWord")) {
  
  if(!is.na(x$main))
    WrdCaption(x$main, wrd = wrd)
  
  font <- rep(font, times=2)[1:2]
  # font[1] is font.txt, font[2] font.desc
  
  ToWrd(x$txt, font = font[1], wrd = wrd)
  ToWrd("\n", wrd = wrd)
  WrdTable(ncol = 2, widths = c(5, 11), wrd = wrd)
  out <- capture.output(x$desc)[-c(1:6)]
  out <- gsub("p-val", "\n  p-val", out)
  out <- gsub("contains", "\n  contains", out)
  ToWrd(out, font = font[2], wrd = wrd)
  wrd[["Selection"]]$MoveRight(wdConst$wdCell, 1, 0)
  WrdPlot(width = 10, height = 6.5, dfact = 2.1, 
          crop = c(0, 0, 0.3, 0), wrd = wrd, append.cr = TRUE)
  
  wrd[["Selection"]]$EndOf(wdConst$wdTable)
  wrd[["Selection"]]$MoveRight(wdConst$wdCharacter, 2, 0)
  wrd[["Selection"]]$TypeParagraph()
  
}




#' @rdname ToWrd
ToWrd.TMod <- function (x, font = NULL, para = NULL, main = NULL, align = NULL, 
                        split=" ", fixed = TRUE, 
                        autofit = TRUE, digits = 3, na.form = "-", ..., 
                        wrd = DescToolsOfficeOptions("lastWord")) {
  
  
  # prepare quality measures  
  x2 <- x[[2]]
  x[[2]][, -1] <- Format(x[[2]][, -1], digits = digits, na.form = na.form)
  x[[2]][x[[2]]$stat %in% c("numdf", "dendf", "N", "n vars", "n coef"), -1] <- 
    Format(x2[x[[2]]$stat %in% c("numdf", "dendf", "N", "n vars", "n coef"), -1], 
           digits = 0, na.form = na.form)
  
  if(!is.null(split)) {
    # xx <- SplitToCol(x[[1]][, -1], split=split, fixed=fixed)
    xx <- SplitToCol(as.data.frame(lapply(x[[1]], StrTrim))[, -1], 
                     split=split, fixed=fixed)
    
    
    zz <- x[[2]][,-1]
    vn <- character()
    for(i in seq_along(attr(xx, "cols"))) {
      j <- attr(xx, "cols")[i]
      zz <- Append(zz, values = matrix("", ncol=j-1), 
                   after = cumsum(c(1, attr(xx, "cols")))[i], names="", stringsAsFactors=FALSE)
      vn <- c(vn, names(attr(xx, "cols"))[i], rep("", j-1))
    }
    
  } else {
    xx <- x[[1]][-1]
    zz <- x[[2]][-1]
  } 
  
  tt <- do.call(rbind, list(SetNames(xx, ""), 
                            SetNames(rep("",  ncol(xx)), ""),
                            SetNames(zz, "")))
  
  ttt <- SetNames(data.frame(c(x[[1]][,1], "---", as.character(x[[2]][,1])), tt, stringsAsFactors = FALSE),
                  c(colnames(x[[1]])[1], vn))
  
  ToWrd(as.matrix(ttt), 
        font=font, 
        align=align)
  
}



