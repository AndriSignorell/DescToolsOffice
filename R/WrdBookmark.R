
#' Get Handle to a MS-Word Bookmark
#' 
#' Accessing bookmarks by name is only possible by browsing the bookmark names.
#' \code{WrdBookmark} returns a handle to a bookmark by taking its name as
#' argument. \code{WrdGotoBookmark} allows to place the
#' cursor on the bookmark.
#' 
#' Bookmarks are useful to build structured documents, which can be updated
#' later.
#' 
#' @aliases WrdBookmark WrdGoto
#' 
#' @param name the name of the bookmark.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOptions("lastWord")}.
#' @param what a word constant, defining the type of object to be used to place
#' the cursor.

#' @author Andri Signorell <andri@@signorell.net>

#' @seealso \code{\link{WrdFont}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}

#' @keywords print

#' @examples
#' 
#' \dontrun{ # we can't get this through the CRAN test - run it with copy/paste to console
#' wrd <- GetNewWrd()
#' WrdText("a)\n\n\nb)", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)
#' WrdInsertBookmark("chap_b")
#' WrdText("\n\n\nc)\n\n\n", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)
#' 
#' WrdGoto("chap_b")
#' WrdUpdateBookmark("chap_b", "Goto chapter B and set text")
#' 
#' WrdInsertBookmark("mybookmark")
#' ToWrd("A longer text\n\n\n")
#' 
#' # Now returning the bookmark
#' bm <- WrdBookmark("mybookmark")
#' 
#' # get the automatically created name of the bookmark
#' bm$name()
#' }

#' @export WrdBookmark


WrdBookmark <- function(name, wrd = DescToolsOptions("lastWord")){
  
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


#' @rdname WrdBookmark
#' @export WrdGoto
WrdGoto <- function (name, what = wdConst$wdGoToBookmark, 
                     wrd = DescToolsOptions("lastWord")) {
  
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




#' Insert New/Delete MS-Word Bookmark
#' \code{WrdInsertBookmark}, \code{WrdDeleteBookmark} inserts/deletes
#' a bookmark in a Word document.
#'  
#' @aliases WrdInsertBookmark WrdDeleteBookmark  
#' 
#' @param name the name of the bookmark.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOptions("lastWord")}.

#' @author Andri Signorell <andri@@signorell.net>

#' @seealso \code{\link{WrdFont}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}

#' @keywords print

#' @examples
#' 
#' \dontrun{ # we can't get this through the CRAN test - run it with copy/paste to console
#' wrd <- GetNewWrd()
#' WrdText("a)\n\n\nb)", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)
#' WrdInsertBookmark("chap_b")
#' WrdText("\n\n\nc)\n\n\n", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)
#' 
#' WrdGoto("chap_b")
#' WrdUpdateBookmark("chap_b", "Goto chapter B and set text")
#' 
#' WrdInsertBookmark("mybookmark")
#' ToWrd("A longer text\n\n\n")
#' 
#' # Now returning the bookmark
#' bm <- WrdBookmark("mybookmark")
#' 
#' # get the automatically created name of the bookmark
#' bm$name()
#' }


#' @export WrdInsertBookmark
WrdInsertBookmark <- function (name, wrd = DescToolsOptions("lastWord")) {
  
  #   With ActiveDocument.Bookmarks
  #   .Add Range:=Selection.Range, Name:="entb"
  #   .DefaultSorting = wdSortByName
  #   .ShowHidden = False
  #   End With
  
  wrdBookmarks <- wrd[["ActiveDocument"]][["Bookmarks"]]
  bookmark <- wrdBookmarks$Add(name)
  invisible(bookmark)
}


#' @rdname WrdInsertBookmark
#' @export WrdDeleteBookmark
WrdDeleteBookmark <- function(name, wrd = DescToolsOptions("lastWord")){
  
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





#' Update MS-Word Bookmark
#' \code{WrdUpdateBookmark} replaces the content
#' within the range of the bookmark in a Word document with the given text.
#' 
#' Bookmarks are useful to build structured documents, which can be updated
#' later.
#' 
#' @aliases WrdUpdateBookmark
#' 
#' @param name the name of the bookmark.
#' @param text the text of the bookmark.
#' @param what a word constant, defining the type of object to be used to place
#' the cursor.
#' @param wrd the pointer to a word instance. Can be a new one, created by
#' \code{GetNewWrd()} or an existing one, created by \code{GetCurrWrd()}.
#' Default is the last created pointer stored in
#' \code{DescToolsOptions("lastWord")}.
#' @author Andri Signorell <andri@@signorell.net>
#' @seealso \code{\link{WrdFont}}, \code{\link{WrdPlot}},
#' \code{\link{GetNewWrd}}, \code{\link{GetCurrWrd}}
#' @keywords print
#' @examples
#' 
#' \dontrun{ # we can't get this through the CRAN test - run it with copy/paste to console
#' wrd <- GetNewWrd()
#' WrdText("a)\n\n\nb)", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)
#' WrdInsertBookmark("chap_b")
#' WrdText("\n\n\nc)\n\n\n", fontname=WrdGetFont()$name, fontsize=WrdGetFont()$size)
#' 
#' WrdGoto("chap_b")
#' WrdUpdateBookmark("chap_b", "Goto chapter B and set text")
#' 
#' WrdInsertBookmark("mybookmark")
#' ToWrd("A longer text\n\n\n")
#' 
#' # Now returning the bookmark
#' bm <- WrdBookmark("mybookmark")
#' 
#' # get the automatically created name of the bookmark
#' bm$name()
#' }
#' @export WrdUpdateBookmark

WrdUpdateBookmark <- function (name, text, what = wdConst$wdGoToBookmark, 
                               wrd = DescToolsOptions("lastWord")) {
  
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








