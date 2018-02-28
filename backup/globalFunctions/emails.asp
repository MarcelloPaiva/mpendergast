<%
	Function send_mail(from, xto, subject, body)
	  '///////////////// send cdo email.

'content of the email
' Send by connecting to port 25 of the SMTP server.
Dim iMsg 
Dim iConf 
Dim Flds 
Dim strHTML

Const cdoSendUsingPort = 2




set iMsg = CreateObject("CDO.Message")
set iConf = CreateObject("CDO.Configuration")

Set Flds = iConf.Fields

' Set the CDOSYS configuration fields to use port 25 on the SMTP server.

With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
    'ToDo: Enter name or IP address of remote SMTP server.
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.marthapendergastrealestate.com" 
'    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.comcast.net" 
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10  
    .Update
End With

' Build HTML for message body.
strHTML = "<HTML>"
strHTML = strHTML & "<HEAD>"
strHTML = strHTML & "<BODY>"
strHTML = strHTML & body
strHTML = strHTML & "</BODY>"
strHTML = strHTML & "</HTML>"


'Set objMessage = Server.CreateObject("CDO.Message") 
iMsg.Subject = subject '"Example CDO Message" 
iMsg.From = from '"martha@marthapendergastrealestate.com" 
iMsg.To = xto
iMsg.HTMLBody = strHTML 
'objMessage.TextBody = "This is some sample message text." 
iMsg.Send 


' Apply the settings to the message.
''With iMsg
  ''  Set .Configuration = iConf
    ''.To = xto 'ToDo: Enter a valid email address.
    ''.From =  from 'ToDo: Enter a valid email address.
    ''.Subject = subject
    ''.HTMLBody = strHTML
    ''.Send
''End With

' Clean up variables.
Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing







'OLD CODE FROM DANILO
'		dim objCDOSYSMail, objCDOSYSCon
'		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
'		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
'		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.marthapendergastrealestate.com"
'		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30 
'		objCDOSYSCon.Fields.update 



'		Set objCDOSYSMail.Configuration = objCDOSYSCon
'		objCDOSYSMail.From = from
'		objCDOSYSMail.To = xto
'		objCDOSYSMail.Subject = subject
'		objCDOSYSMail.htmlBody = body
'		objCDOSYSMail.Send 
'		Set objCDOSYSMail = Nothing 
'		Set objCDOSYSCon = Nothing 
		
		
		
	end function
%>