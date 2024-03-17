


## PowerPoint functions ====



#' Add Slides, Insert Texts and Plots to PowerPoint
#' 
#' A couple of functions to get R-stuff into MS-Powerpoint.
#' 
#' \code{GetNewPP()} starts a new instance of PowerPoint and returns its
#' handle. A new presentation with one empty slide will be created thereby. The
#' handle is needed for addressing the presentation afterwards.\cr
#' \code{GetCurrPP()} will look for a running PowerPoint instance and return
#' its handle. \code{NULL} is returned if nothing's found. \code{PpAddSlide()}
#' inserts a new slide into the active presentation.\cr \code{PpPlot()} inserts
#' the active plot into PowerPoint. The image is transferred by saving the
#' picture to a file in R and inserting the file in PowerPoint. The format of
#' the plot can be selected, as well as crop options and the size factor for
#' inserting.\cr \code{PpText()} inserts a new textbox with given text and box
#' properties.
#' 
#' See PowerPoint-objectmodel for further informations. %% ~~ If necessary,
#' more details than the description above ~~
#' 

#' @aliases PpPlot PpText PpAddSlide

#' @param pos position of the new inserted slide within the presentation.
#' @param type the format for the picture file, default is \code{"png"}.
#' @param crop crop options for the picture, defined by a 4-elements-vector.
#' The first element is the bottom side, the second the left and so on.
#' @param picscale scale factor of the picture in percent, default ist 100.
#' @param x,y left/upper xy-coordinate for the plot or for the textbox.
#' @param height height in cm, this overrides the picscale if both are given.
#' @param width width in cm, this overrides the picscale if both are given.
#' @param res resolution for the png file, defaults to 200.
#' @param dfact the size factor for the graphic.
#' @param txt text to be placed in the textbox
#' @param fontname used font for textbox
#' @param fontsize used fontsize for textbox
#' @param bold logic. Text is set bold if this is set to \code{TRUE} (default
#' is FALSE).
#' @param italic logic. Text is set italic if this is to \code{TRUE} (default
#' is FALSE).
#' @param col font color, defaults to \code{"black"}.
#' @param bg background color for textboxdefaults to \code{"white"}.
#' @param hasFrame logical. Defines if a textbox is to be framed. Default is
#' TRUE.
#' @param pp the pointer to a PowerPoint instance, can be a new one, created by
#' \code{GetNewPP()} or the last created by \code{DescToolsOptions("lastPP")}
#' (default).
#' @return The functions return the pointer to the created object.

#' @author Andri Signorell <andri@@signorell.net>

#' @seealso \code{\link{WrdPlot}} 
#' @keywords print
#' @examples
#' 
#' \dontrun{# Windows-specific example
#' 
#' # let's have some graphic
#' plot(1,type="n", axes=FALSE, xlab="", ylab="", xlim=c(0,1), ylim=c(0,1))
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
#' # get a handle to a new PowerPoint instance
#' pp <- GetNewPP()
#' # insert plot with a specified height
#' PpPlot(pp=pp,  x=150, y=150, height=10, width=10)
#' 
#' PpText("Remember?\n", fontname="Arial", x=200, y=70, height=30, fontsize=14,
#'        bold=TRUE, pp=pp, bg="lemonchiffon", hasFrame=TRUE)
#' 
#' PpAddSlide(pp=pp)
#' # crop the picture
#' pic <- PpPlot(pp=pp, x=1, y=200, height=10, width=10, crop=c(9,9,0,0))
#' pic
#' 
#' 
#' # some more automatic procedure
#' pp <- GetNewPP()
#' PpText("Hello to my presentation", x=100, y=100, fontsize=32, bold=TRUE,
#'        width=300, hasFrame=FALSE, col="blue", pp=pp)
#' 
#' for(i in 1:4){
#'   barplot(1:4, col=i)
#'   PpAddSlide(pp=pp)
#'   PpPlot(height=15, width=21, x=50, y=50, pp=pp)
#'   PpText(gettextf("This is my barplot nr %s", i), x=100, y=10, width=300, pp=pp)
#' }
#' }
#' 


PpAddSlide <- function(pos = NULL, pp = DescToolsOptions("lastPP", default = GetNewPP())){
  
  slides <- pp[["ActivePresentation"]][["Slides"]]
  if(is.null(pos)) pos <- slides$Count()+1
  slides$AddSlide(pos, slides$Item(1)[["CustomLayout"]])$Select()
  
  invisible()
}

#' @rdname PpAddSlide
PpText <- function (txt, x=1, y=1, height=50, width=100, fontname = "Calibri", fontsize = 18, 
                    bold = FALSE,
                    italic = FALSE, col = "black", bg = "white", hasFrame = TRUE, 
                    pp = DescToolsOptions("lastPP", default = GetNewPP())) {
  
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


#' @rdname PpAddSlide
PpPlot <- function( type="png", crop=c(0,0,0,0),
                    picscale=100, x=1, y=1, height=NA, width=NA, res=200, dfact=1.6, 
                    pp = DescToolsOptions("lastPP", default = GetNewPP()) ){
  
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


