Set Batch = ChronoApp.GetCurrentBatch
CurrentBatchname=Batch.Getname
Dim NumDocs
Dim errDocs
NumDocs=Batch.GetDocCount

'navigating records to find number of error docs
errDocs=0
For numDoc = 0 To NumDocs-1
	If(Batch.IsValidated(numDoc))=0 Then
	errDocs=errDocs+1
	FileName=Batch.GetSystemField(numDoc ,"SrcDoc")'Get image name
	InvNoValstatus=Batch.GetSystemField(numDoc ,"sysval_Invoice Number")'Get invoice no validation status
	
	ErrorString=Errorstring &   vbcrlf&_  
	"Document No = " & numDoc +1 & " Batch=" & CurrentBatchname & vbcrlf&_
	"Error File = " & FileName & vbcrlf&_
	"Invoice Number Field Status = " & InvNoValstatus & vbcrlf
	
	
'msgbox ErrorString
	
Else
End If
Next

If errDocs > 0 Then

strSMTPFrom = "joe@joe.com"
strSMTPTo = "bob@gmail.com"
strSMTPRelay = "myrelay.domain.net"
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

	
