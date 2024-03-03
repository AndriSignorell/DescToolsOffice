


# ToWrd functions



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



