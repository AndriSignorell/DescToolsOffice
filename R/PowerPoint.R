


## PowerPoint functions ====





PpAddSlide <- function(pos = NULL, pp = getOption("lastPP", default = GetNewPP())){
  
  slides <- pp[["ActivePresentation"]][["Slides"]]
  if(is.null(pos)) pos <- slides$Count()+1
  slides$AddSlide(pos, slides$Item(1)[["CustomLayout"]])$Select()
  
  invisible()
}



PpText <- function (txt, x=1, y=1, height=50, width=100, fontname = "Calibri", fontsize = 18, 
                    bold = FALSE,
                    italic = FALSE, col = "black", bg = "white", hasFrame = TRUE, 
                    pp = getOption("lastPP", default = GetNewPP())) {
  
  msoShapeRectangle <- 1
  
  if (!inherits(x=txt, what="character"))
    txt <- DescTools:::.CaptOut(txt)
  #  slide <- pp[["ActivePresentation"]][["Slides"]]$Item(1)
  slide <- pp$ActiveWindow()$View()$Slide()
  shape <- slide[["Shapes"]]$AddShape(msoShapeRectangle, x, y, x + width, y+height)
  textbox <- shape[["TextFrame"]]
  textbox[["TextRange"]][["Text"]] <- txt
  
  tbfont <- textbox[["TextRange"]][["Font"]]
  tbfont[["Name"]] <- fontname
  tbfont[["Size"]] <- fontsize
  tbfont[["Bold"]] <- bold
  tbfont[["Italic"]] <- italic
  tbfont[["Color"]] <- RgbToLong(ColToRgb(col))
  
  textbox[["MarginBottom"]] <- 10
  textbox[["MarginLeft"]] <- 10
  textbox[["MarginRight"]] <- 10
  textbox[["MarginTop"]] <- 10
  
  shp <- shape[["Fill"]][["ForeColor"]]
  shp[["RGB"]] <- RgbToLong(ColToRgb(bg))
  shp <- shape[["Line"]]
  shp[["Visible"]] <- hasFrame
  
  invisible(shape)
  
}





PpPlot <- function( type="png", crop=c(0,0,0,0),
                    picscale=100, x=1, y=1, height=NA, width=NA, res=200, dfact=1.6, 
                    pp = getOption("lastPP", default = GetNewPP()) ){
  
  # height, width in cm!
  # scale will be overidden, if height/width defined
  
  # Example: PpPlot(picscale=30)
  #          PpPlot(width=8)
  
  CmToPts <- function(x) x * 28.35
  PtsToCm <- function(x) x / 28.35
  # http://msdn.microsoft.com/en-us/library/bb214076(v=office.12).aspx
  
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
  
  
  # slide <- pp[["ActivePresentation"]][["Slides"]]$Item(1)
  slide <- pp$ActiveWindow()$View()$Slide()
  pic <- slide[["Shapes"]]$AddPicture(fn, FALSE, TRUE, x, y)
  
  picfrmt <- pic[["PictureFormat"]]
  picfrmt[["CropBottom"]] <- CmToPts(crop[1])
  picfrmt[["CropLeft"]] <- CmToPts(crop[2])
  picfrmt[["CropTop"]] <- CmToPts(crop[3])
  picfrmt[["CropRight"]] <- CmToPts(crop[4])
  
  if( is.na(height) & is.na(width) ){
    # or use the ScaleHeight/ScaleWidth attributes:
    msoTrue <- -1
    msoFalse <- 0
    pic$ScaleHeight(picscale/100, msoTrue)
    pic$ScaleWidth(picscale/100, msoTrue)
    
  } else {
    # Set new height:
    if( is.na(width) ) width <- height / PtsToCm( pic[["Height"]] ) * PtsToCm( pic[["Width"]] )
    if( is.na(height) ) height <- width / PtsToCm( pic[["Width"]] ) * PtsToCm( pic[["Height"]] )
    pic[["Height"]] <- CmToPts(height)
    pic[["Width"]] <- CmToPts(width)
  }
  
  if( file.exists(fn) ) { file.remove(fn) }
  
  invisible( pic )
  
}


