Set Batch = ChronoApp.GetCurrentBatch
Dim NumDocs
Dim errDocs
NumDocs=Batch.GetDocCount
 
'navigating records to find number of error docs
errDocs=0
For numDoc = 0 To NumDocs-1
    If(Batch.IsValidated(numDoc))=0 Then
        errDocs=errDocs+1
    Else
    End If
Next

If errDocs > 0 Then

    strSMTPFrom = "no-reply@yourcompany.com"
    strSMTPTo = "joe@gmail.com"
    strSMTPRelay = "relay.isp.net"
    strTextBody = "Number of errors = " & errDocs
    strSubject = "Subject line"
    'strAttachment = "c:\this_attachment.pdf"


    Set oMessage = CreateObject("CDO.Message")
    oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
    oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPRelay
    oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
    oMessage.Configuration.Fields.Update

    oMessage.Subject = strSubject
    oMessage.From = strSMTPFrom
    oMessage.To = strSMTPTo
    oMessage.TextBody = strTextBody
    'oMessage.AddAttachment strAttachment


    oMessage.Send

    Else
End If
 
