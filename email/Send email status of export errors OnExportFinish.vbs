Set Batch = ChronoApp.GetCurrentBatch
CurrentBatchname=Batch.Getname
Dim NumDocs
Dim errDocs
NumDocs=Batch.GetDocCount

'navigating records to find number of error docs
errDocs=0

For numDoc = 0 To NumDocs-1
    Set Doc=Batch.GetDocument(numDoc)
    
    ' You will need to customise the string "expinfo.exported.ChronoScan PDF Text Conversion.filesystem1" to match your export connector
    expStatus=Doc.get_field_value("expinfo.exported.ChronoScan PDF Text Conversion.filesystem1")
   
    If expStatus="0" Then
        errDocs=errDocs+1
    End If

    FileName=Batch.GetSystemField(numDoc ,"SrcDoc")'Get image name
  
    ErrorString=Errorstring &   vbcrlf&_  
    "Document No = " & numDoc +1 & " Batch=" & CurrentBatchname & vbcrlf&_
    "Error File = " & FileName & vbcrlf&_
    "Invoice Number Field Status = " & InvNoValstatus & vbcrlf
    
    
'msgbox ErrorString
    

Next

If errDocs > 0 Then

strSMTPFrom = "joe@joe.com"
strSMTPTo = "pafowkes@gmail.com"
strSMTPRelay = "relay.plus.net"
strTextBody =   "Number of errors = " & errDocs & vbcrlf&_
ErrorString

strSubject = "Validation Errors"
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
