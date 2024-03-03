

printWrd <- function(x, main = NULL, plotit = NULL, ..., wrd = wrd) {
  # x is a Desc object, wrd the handle to a word instance
  
  WrdPlotDesc <- function(z, wrd) {
    .plotReset <- function() {
      layout(matrix(1))
      par(
        xlog = FALSE, ylog = FALSE, adj = 0.5, ann = TRUE,
        ask = FALSE, bg = "white", bty = "o", cex = 1, cex.axis = 1,
        cex.lab = 1, cex.main = 1.2, cex.sub = 1, col = "black",
        col.axis = "black", col.lab = "black", col.main = "black",
        col.sub = "black", crt = 0, err = 0L, family = "", fg = "black",
        fig = c(0, 1, 0, 1), fin = c(12.8333333333333, 8), font = 1L,
        font.axis = 1L, font.lab = 1L, font.main = 2L, font.sub = 1L,
        #      lab = c(5L, 5L, 7L), las = 0L, lend = "round", lheight = 1,
        lab = c(5L, 5L, 7L), lend = "round", lheight = 1,
        ljoin = "round", lmitre = 10, lty = "solid", lwd = 1,
        mai = c(1.36, 1.09333, 1.093333, 0.56), mar = c(5.1, 4.1, 4.1, 2.1),
        mex = 1, mfcol = c(1L, 1L), mfg = c(1L, 1L, 1L, 1L),
        mfrow = c(1L, 1L), mgp = c(3, 1, 0), mkh = 0.001, new = FALSE,
        oma = c(0, 0, 0, 0), omd = c(0, 1, 0, 1), omi = c(0, 0, 0, 0),
        pch = 1L, pin = c(11.18, 5.54666666666667),
        plt = c(0.0851948051948052, 0.956363636363636, 0.17, 0.863333333333333),
        ps = 16L, pty = "m", smo = 1, srt = 0, tck = NA_real_,
        tcl = -0.5, usr = c(0, 1, 0, 1), xaxp = c(0, 1, 5),
        xaxs = "r", xaxt = "s", xpd = FALSE,
        yaxp = c(0, 1, 5), yaxs = "r", yaxt = "s", ylbias = 0.2
      )
      #   par(
      #     xlog = FALSE, ylog = FALSE,
      #     mai = c(1.36, 1.09333, 1.093333, 0.56), mar = c(5.1, 4.1,4.1, 2.1),
      #     mex = 1, mfcol = c(1L, 1L), mfg = c(1L, 1L, 1L, 1L),
      #     mfrow = c(1L, 1L),
      #     oma = c(0, 0, 0, 0), omd = c(0, 1, 0, 1), omi = c(0, 0, 0, 0),
      #     usr = c(0, 1, 0, 1), xpd = FALSE
      #     )
    }
    
    
    .plotReset()
    
    if (identical(z[[1]]$noplot, TRUE)) {
      # identical as noplot will not be present in filled objects!!
      # there's nothing to plot, the variable might be empty, so just leave here
    } else {
      if (any(z[[1]]$class %in% c("factor", "ordered", "character") ||
              (z[[1]]$class == "integer" && !is.null(z[[1]]$freq)))) {
        plot(z, main = NA)
        WrdPlot(
          width = 8, height = pmin(2 + 3 / 6 * nrow(z[[1]]$freq), 10),
          dfact = 2.7, crop = c(0, 0, 0, 0), wrd = wrd, append.cr = FALSE
        )
      } else if (any(z[[1]]$class %in% c("numeric", "integer"))) {
        plot(z, main = NA)
        WrdPlot(
          width = 8, height = 5.0, dfact = 2.3,
          crop = c(-.2, 0, 0, 0), wrd = wrd, append.cr = FALSE
        )
      } else if (any(z[[1]]$class %in% "logical")) {
        plot(z, main = NA)
        WrdPlot(
          width = 6, height = 4, dfact = 2.6,
          crop = c(-.2, 0.2, 1, 0), wrd = wrd, append.cr = FALSE
        )
      } else if (z[[1]]$class == "Date") {
        plot(z, main = NA, type = 1)
        WrdPlot(
          width = 6.5, height = 5, dfact = 2.5, wrd = wrd,
          append.cr = TRUE
        )
        plot(z, main = NA, type = 2)
        WrdPlot(
          width = 6.5, height = 6.2, dfact = 2.5, wrd = wrd,
          append.cr = TRUE
        )
        plot(z, main = NA, type = 3)
        WrdPlot(
          width = 6.5, height = 4, dfact = 2.5, wrd = wrd,
          append.cr = TRUE
        )
      } else if (z[[1]]$class %in% c("table", "matrix", "factfact")) {
        plot(z, main = NA, horiz = z[[1]]$horiz)
        if (z[[1]]$horiz) {
          WrdPlot(
            width = 16, height = 6.5, dfact = 2.5, wrd = wrd,
            append.cr = TRUE
          )
        } else {
          WrdPlot(
            width = 7, height = 14, dfact = 2.5, wrd = wrd,
            append.cr = TRUE
          )
        }
      } else if (z[[1]]$class %in% c("numnum")) {
        plot(z, main = NA)
        WrdPlot(
          width = 6.5, height = 6.5 / DescTools:::gold_sec_c, dfact = 2.5,
          crop = c(0, 0, 0.2, 0), wrd = wrd, append.cr = TRUE
        )
      } else if (z[[1]]$class %in% c("numfact")) {
        plot(z, main = NA)
        WrdPlot(
          width = 15, height = 7, dfact = 2.2,
          crop = c(0, 0, 0.2, 0), wrd = wrd, append.cr = TRUE
        )
      } else if (z[[1]]$class %in% c("factnum")) {
        plot(z, main = NA)
        WrdPlot(
          width = 15, height = 7, dfact = 2.2,
          crop = c(0, 0, 0.2, 0), wrd = wrd, append.cr = TRUE
        )
      }
    }
    invisible()
  }
  
  
  # start main proc  ****************
  
  # get fixed font
  fixedfont <- getOption("fixedfont", list(name = "Consolas", size = 7))
  
  for (i in seq_along(x)) {
    # # skip object header entries
    # if(names(x[i]) == "_objheader")
    #   next
    
    if (x[[i]]$class == "header") {
      if (is.null(x[[i]][["abstract"]])) {
        txt <- DescTools:::.CaptOut(print(x[i]))[-(1:2)]
        WrdCaption(x[[i]]$main, wrd = wrd)
        ToWrd(txt = txt, wrd = wrd)
        # WrdText(txt=txt, wrd=wrd )
      } else {
        attr(x[[i]]$abstract, "main") <- x[[i]][["main"]]
        ToWrd(x[[i]]$abstract, wrd = wrd)
      }
    } else {
      WrdCaption(x[[i]]$main, wrd = wrd)
      
      if (!is.null(x[[i]]$label)) {
        lblfont <- InDots(..., arg = "font", default = list(size = 8))
        lblfont$size <- 8
        ToWrd.character(
          x = paste("\n", x[[i]]$label, "\n", sep = ""),
          font = lblfont, wrd = wrd
        )
      }
      
      
      txt <- DescTools:::.CaptOut(print(x[i], nolabel = TRUE))[-(1:2)]
      
      if (x[[i]]$class == "Date") {
        WrdTable(nrow = 4, ncol = 2, wrd = wrd)
        # merge cells in the first row
        wrd[["Selection"]]$MoveRight(
          Unit = wdConst$wdCharacter, Count = 2,
          Extend = wdConst$wdExtend
        )
        wrd[["Selection"]][["Cells"]]$Merge()
        
        ToWrd(x = txt[1:6], font = fixedfont, wrd = wrd)
        wrd[["Selection"]]$MoveRight(wdConst$wdCell, 1, 0)
        ToWrd(x = txt[-c(1:6)], font = fixedfont, wrd = wrd)
      } else {
        if (max(unlist(lapply(txt, nchar))) < 59) {
          # decide if two rows or 2 columns ist adequate
          WrdTable(nrow = 1, ncol = 2, wrd = wrd)
          x[[i]]$horiz <- FALSE
        } else {
          WrdTable(nrow = 2, ncol = 1, wrd = wrd)
          x[[i]]$horiz <- TRUE
        }
        
        ToWrd(x = txt, font = fixedfont, wrd = wrd)
      }
      
      wrd[["Selection"]]$MoveRight(wdConst$wdCell, 1, 0)
      
      plotit <- Coalesce(plotit, x$plotit, DescToolsOptions("plotit"), FALSE)
      if (plotit) {
        WrdPlotDesc(x[i], wrd = wrd)
      }
      
      
      wrd[["Selection"]]$EndOf(wdConst$wdTable)
      # get out of tablerange
      wrd[["Selection"]]$MoveRight(wdConst$wdCharacter, 2, 0)
      selborder <- wrd[["Selection"]]$Borders(wdConst$wdBorderTop)
      selborder[["LineStyle"]] <- wdConst$wdLineStyleSingle
      wrd[["Selection"]]$TypeParagraph()
    }
  }
  
  invisible()
}

