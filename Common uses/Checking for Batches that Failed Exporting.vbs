' https://www.chronoscan.org/doc/checking_for_batches_that_failed_exporting.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

' This sample can be used to check for Batches that have errors. This script will generate a message that is sent via email if there are any batches on the system that returned errors while exporting.
'Just replace "JOB NAME" with the desired Job name or leave it empty to loop through all existing batches.
 
batchesArray = ChronoApp.GetBatches("JOB NAME")
 
BatchStatArr = Array("not exported","exported ok","not exported because of not validated documents","error exporting","export canceled")
 
strMessage = ""
 
For Each strBatch In batchesArray
    Set Batch = ChronoApp.CreateBatch("JOB NAME",strBatch)
    batchStatus = Batch.GetExportStatus
    If batchStatus = 1 Then
        strMessage = strMessage
    Else
    strMessage = strMessage & "Batch Name: " & strBatch & " Batch Status: " & BatchStatArr(batchStatus)
    End If
Next
 
   
   
 
If strMessage <> "" Then
'msgbox strMessage
 
 
'The code for the email is the following. Several details can be added including attachments
 
    strSMTPFrom = "TOemail@domain.com"
    strSMTPTo = "FROMemail@somedomain.com"
    strSMTPRelay = "yourrelay.domain.net"
    strTextBody = ""
    strSubject = BatchName & " has " & NumErrors & " validation error(s)"
 
    'A good idea is to include a file with the export report as an attachment
 
    'strAttachment = "c:\this_attachment.pdf"
 
 
    Set oMessage = CreateObject("CDO.Message")
    oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPRelay
    oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    oMessage.Configuration.Fields.Update
 
    oMessage.Subject = strSubject
    oMessage.From = strSMTPFrom
    oMessage.To = strSMTPTo
    oMessage.TextBody = strMessage
    'oMessage.AddAttachment strAttachment
 
 
    oMessage.Send
 
Else
End If

