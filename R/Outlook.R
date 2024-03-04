# Outlook




#' Send a Mail Using Outlook as Mail Client
#' 
#' Sending emails in R can be required in some reporting tasks. As we already
#' have RDCOMClient available we wrap the send code in a function.
#' 
#' 
#' @param to a vector of recipients
#' @param cc a vector of recipients receiving a carbon copy
#' @param bcc a vector of recipients receiving a blind carbon copy
#' @param subject the subject of the mail
#' @param body the body text of the mail
#' @param attachment a vector of paths to attachments
#' @return Nothing is returned
#' @author Andri Signorell <andri@@signorell.net> strongly based on code of
#' Franziska Mueller
#' @seealso \code{\link{ToXL}}
#' @keywords MS-Office
#' @examples
#' \dontrun{
#' SendOutlookMail(to=c("me@microsoft.com", "you@rstudio.com"), subject = "Some Info", 
#'                 body = "Hi all\r Find the files attached\r Regards, Andri", 
#'                 attachment = c("C:/temp/fileA.txt", 
#'                                "C:/temp/fileB.txt"))
#' }
#' 
#' @export SendOutlookMail
SendOutlookMail <- function(to, cc=NULL, bcc=NULL, subject, body, attachment=NULL){
  
  out <- GetCOMAppHandle("Outlook.Application", existing=TRUE)
  
  mail <- out$CreateItem(0)
  mail[["to"]] <- to
  if(!is.null(cc)) mail[["cc"]] <- cc
  if(!is.null(bcc)) mail[["bcc"]] <- bcc
  mail[["subject"]] <- subject
  mail[["body"]] <- body
  
  ## Add attachments
  if(!is.null(attachment)) 
    sapply(attachment, function(x) mail[["Attachments"]]$Add(x))
  
  ## senden                  
  mail$Send()
  
  rm(out, mail)
  gc() 
  
  invisible()
  
}


