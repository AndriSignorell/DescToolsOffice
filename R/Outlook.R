

# Outlook


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


