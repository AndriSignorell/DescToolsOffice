





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







ToWrd.default <- function(x, font=NULL, ..., wrd=DescToolsOfficeOptions("lastWord")){
  
  ToWrd.character(x=DescTools:::.CaptOut(x), font=font, ..., wrd=wrd)
  invisible()
  
}



ToWrd.Desc <- function(x, font=NULL, ..., wrd=DescToolsOfficeOptions("lastWord")){
  
  printWrd(x, ..., wrd=wrd)
  invisible()
  
}




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


WrdCaption <- function(x, index = 1, wrd = DescToolsOfficeOptions("lastWord")){
  
  lst <- Recycle(x=x, index=index)
  x <-
    index <- lst[["index"]]
  for(i in seq(attr(lst, "maxdim")))
    ToWrd.character(paste(lst[["x"]][i], "\n", sep = ""),
                    style = eval(parse(text = gettextf("wdConst$wdStyleHeading%s", lst[["index"]][i]))))
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






WrdCellRange <- function(wtab, from, to) {
  # returns a handle for the table range
  wtrange <- wtab[["Parent"]]$Range(
    wtab$Cell(from[1], from[2])[["Range"]][["Start"]],
    wtab$Cell(to[1], to[2])[["Range"]][["End"]]
  )
  
  return(wtrange)
}


WrdMergeCells <- function(wtab, rstart, rend) {
  
  rng <- WrdCellRange(wtab, rstart, rend)
  rng[["Cells"]]$Merge()
  
}

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


WrdStyle <- function (wrd = DescToolsOfficeOptions("lastWord")) {
  wrdSel <- wrd[["Selection"]]
  wrdStyle <- wrdSel[["Style"]][["NameLocal"]]
  return(wrdStyle)
}


`WrdStyle<-` <- function (wrd, value) {
  wrdSel <- wrd[["Selection"]][["Paragraphs"]]
  wrdSel[["Style"]] <- value
  return(wrd)
}




WrdGoto <- function (name, what = wdConst$wdGoToBookmark, wrd = DescToolsOfficeOptions("lastWord")) {
  wrdSel <- wrd[["Selection"]]
  
  if(what == wdConst$wdGoToBookmark){
    wrdBookmarks <- wrd[["ActiveDocument"]][["Bookmarks"]]
    if(wrdBookmarks$exists(name)){
      wrdSel$GoTo(what=what, Name=name)
      res <- TRUE
    } else {
      warning(gettextf("Bookmark %s does not exist, so there's nothing to select", name))
      res <- FALSE
    }
  } else {
    wrdSel$GoTo(what=what, Name=name)
    
  }
  
  invisible()
}






WrdPageBreak <- function(wrd = DescToolsOfficeOptions("lastWord")) {
  wrd[["Selection"]]$InsertBreak(wdConst$wdSectionBreakNextPage)
  invisible()
}



WrdBookmark <- function(name, wrd = DescToolsOfficeOptions("lastWord")){
  
  wbms <- wrd[["ActiveDocument"]][["Bookmarks"]]
  
  if(wbms$count()>0){
    # get bookmark names
    bmnames <- sapply(seq(wbms$count()), function(i) wbms[[i]]$name())
    
    id <- which(name == bmnames)
    
    if(length(id)==0)   # name found?
      res <- NULL 
    
    else
      res <- wbms[[id]]
    # no attributes for S4 objects... :-(
    #  res@idx <- which(name == bmnames)
    
  } else {
    # warning(gettextf("bookmark %s not found", bookmark))
    res <- NULL
  }
  
  return(res)  
  
}


# WrdGetBookmarkID <- function(name, wrd = DescToolsOfficeOptions("lastWord")){
#   
#   wrdBookmarks <- wrd[["ActiveDocument"]][["Bookmarks"]]
#   
#   if(wrdBookmarks$exists(name)){
#     if((n <- wrdBookmarks$count()) > 0) {
#       for(i in 1:n){
#         if(name == wrdBookmarks[[i]]$name())
#           return(i)
#       }
#     }
#   } else {
#     warning(gettextf("Bookmark %s does not exist.", name))
#     return(NA_integer_)
#   }
#   
# }




WrdInsertBookmark <- function (name, wrd = DescToolsOfficeOptions("lastWord")) {
  
  #   With ActiveDocument.Bookmarks
  #   .Add Range:=Selection.Range, Name:="entb"
  #   .DefaultSorting = wdSortByName
  #   .ShowHidden = False
  #   End With
  
  wrdBookmarks <- wrd[["ActiveDocument"]][["Bookmarks"]]
  bookmark <- wrdBookmarks$Add(name)
  invisible(bookmark)
}


WrdUpdateBookmark <- function (name, text, what = wdConst$wdGoToBookmark, wrd = DescToolsOfficeOptions("lastWord")) {
  
  #   With ActiveDocument.Bookmarks
  #   .Add Range:=Selection.Range, Name:="entb"
  #   .DefaultSorting = wdSortByName
  #   .ShowHidden = False
  #   End With
  
  wrdSel <- wrd[["Selection"]]
  wrdSel$GoTo(What=what, Name=name)
  wrdSel[["Text"]] <- text
  # the bookmark will be deleted, how can we avoid that?
  wrdBookmarks <- wrd[["ActiveDocument"]][["Bookmarks"]]
  wrdBookmarks$Add(name)
  invisible()
}




WrdDeleteBookmark <- function(name, wrd = DescToolsOfficeOptions("lastWord")){
  
  wrdBookmarks <- wrd[["ActiveDocument"]][["Bookmarks"]]
  if(wrdBookmarks$exists(name)){
    WrdBookmark(name)$Delete()
    res <- TRUE
  } else {
    warning(gettextf("Bookmark %s does not exist, so there's nothing to delete", name))
    res <- FALSE
  }
  
  return(res)
  # TRUE for success / FALSE for fail
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



###


WrdKill <- function(){
  # Word might not always quit and end the task
  # so killing the task is "ultima ratio"...
  
  shell('taskkill /F /IM WINWORD.EXE')
}





###


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

