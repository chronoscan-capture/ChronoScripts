Dim HtmlBody 'To store html body 
Dim EmailSubject 'To store email Subject
Dim EmailFrom 'To store email From
Dim EmailTo 'To store email To
Dim EmailBcc 'To store email BCC
Dim EmailCc 'To store email CC
Dim EmailSMTPServer 'To store SMTP Server name or address
Dim fso 'To declare File System Object for file access
Dim filename 'To declare filename for email content
Dim myattachment 'To declare filename for attachment
 
'Initialization of Email Configuration
EmailSMTPServer = "smtp.office365.com"
EmailSubject = "Test 250717"
EmailFrom = "joe@joe.com"
EmailTo = "bob@gmail.com"
EmailBcc = ""
EmailCc = ""
 
'declare file to use as HTML email content
filename = "C:\temp\myfile.html"
 
'declare file to use as attachment
myattachment = "c:\temp\file1.jpg"
 
'Creating File System Object and loading HTML file into email content
Set fso = CreateObject("Scripting.FileSystemObject")
Set ObjOutFile = fso.OpenTextFile(filename,1)
HtmlBody = objOutFile.ReadAll
 
'Send the HTML email
SendEmail(HtmlBody)
 
'Specify Email sending properties 
Function GetCDOFromProperties() 
Dim SMTPServer, SMTPPort, SMTPAuthenticate, SMTPUserName, SMTPPassword, SMTPTimeout, SMTPProxyServer, SMTPProxyBypass 
SMTPServer = "smtp.office365.com"
SMTPPort = "25" 
SMTPAuthenticate = "1" 
SMTPUserName = ""
SMTPPassword = ""
SMTPTimeout = "50"
SMTPProxyServer = ""
SMTPProxyBypass = ""

 

Set GetCDOFromProperties = ConfigureCDO(SMTPServer, SMTPPort, SMTPAuthenticate, SMTPUserName, SMTPPassword, SMTPTimeout, SMTPProxyServer, SMTPProxyBypass)

End Function

'Configure CDO Email
Function ConfigureCDO(SMTPServer, SMTPPort, SMTPAuthenticate, SMTPUserName, SMTPPassWord, SMTPTimeout, SMTPProxyServer, SMTPProxyBypass)

Dim Conf

' Setup SMTP Options
Set Conf = CreateObject("CDO.Configuration") 

'Specifies the method used to send messages. 1 - Pickup; 2 - Port. If a local SMTP service is available, this field defaults to 1
Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

'   The pickup directory for the local SMTP service. This value is set automatically when the SMTP service is installed.
'Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\mailroot\pickup"

'   The name (DNS) or IP address of the machine hosting the SMTP service through which messages are to be sent.
Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

'   The port on which the SMTP service specified by the smtpserver field is listening for connections. The default and well-known port for an SMTP service is 25.
Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

'   SMTP is using SSL; Default False
Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true

'   0 - Do not auth; 1 - Basic (clear text auth); 2 - Use NTLM authentication (Secure Password Authentication in Microsoft® Outlook® Express) The current process security context is used to authenticate with the service.
Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = SMTPAuthenticate 

Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPUserName
Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassWord


'   Indicates the number of seconds to wait for a valid socket to be established with the SMTP service before timing out. Default is 30
Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SMTPTimeout 

'   The proxy server to use when accessing HTTP resources. Both the name (or IP address) of the proxy machine and the port number must be specified using the format "servername:port."
'Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlproxyserver") = strnullcheck(SMTPProxyServer) '"server:80"

'   Used to specify that for local addresses, the proxy (set with urlproxyserver) should be bypassed. The value of the string, if set, should only be "<local>". This is the same value used with Internet Explorer when setting a proxy with bypass for local addresses. 
'Conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/urlproxybypass") = strnullcheck(SMTPProxyBypass) '"<local>"

Conf.Fields.Update

Set ConfigureCDO = Conf
End Function

'Used to check for blank values in CDO Configuration
Function strnullcheck(varString)
If Isnull(varString) Then
strnullcheck = ""
Else
strnullcheck = varString
End If 
End Function

'Send the email (including an attachment)
Function SendEmail(HtmlBody)
Set objMessage = CreateObject("CDO.Message")
objMessage.Configuration = GetCDOFromProperties()

'Email header 
objMessage.Subject = EmailSubject
objMessage.From = EmailFrom
objMessage.To = EmailTo
objMessage.Bcc = EmailBcc
objMessage.Cc = EmailCc

'Add Email Attachment
objMessage.AddAttachment myattachment

'The line below shows how to send using HTML included directly in your script
objMessage.HTMLBody = HtmlBody

objMessage.Send

Set objMessage = nothing
HtmlBody = "" 
End Function

      
