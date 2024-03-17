
# **************************************
# Some Outlook-Code
# **************************************


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
#' 
#' @return \code{TRUE} is returned in case of success and \code{FALSE} in case of an
#'  unhandled error while sending mail
#' 
#' @author Andri Signorell <andri@@signorell.net> strongly based on code of
#' Franziska Mueller
#' 
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
  
  # Outlook has no visible property
  out <- RDCOMClient::COMCreate("Outlook.Application", force=TRUE, existing=FALSE)

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
  res <- mail$Send()
  
  rm(out, mail)
  gc() 
  
  return(res)
  
  # if(res)
  #   return("Mail sent.")
  # else 
  #   return("Unknown error while sending mail. Please check!")
  
}


