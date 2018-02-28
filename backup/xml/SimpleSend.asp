<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>Chilkat WebMail Example Program - Send a Simple Plain-Text Email</TITLE>
</HEAD>
<BODY>

<%
	' Create a mailman
	set mailman = Server.CreateObject("ChilkatWebMail.WebMailMan")

	' Unlock the component - use your trial or purchased unlock code.
	mailman.UnlockComponent "unlock_code"

	' Tell the mailman where the SMTP server is located
	mailman.SmtpHost = "mail.earthlink.net"
	' Some SMTP servers require a login/password
	'mailman.SmtpUsername = "myLogin"
	'mailman.SmtpPassword = "myPassword"
	
	' Create an Email message
	set email = Server.CreateObject("ChilkatWebMail.WebEmail")

	' Enter the recipient's information
	email.AddTo "John Smith", "jsmith@chilkatsoft.com"

	' Enter the sender's information
	email.FromName = "Joe Smith"
	email.FromAddress = "joe.smith@somedomain.com"

	' Enter the email subject
	email.Subject = "Simple plain-text email"

	email.Body = "This is a test"

	' Sends the email with the patterns replaced.
	if mailman.SendEmail(email) then
		Response.write "Message sent successfully!<br><br>"
	else
		Response.write "ERROR: Message not sent!<br><br>"
	end if
%>
    <br><b>HTML Log:</b><br>
<%
    Response.write mailman.ErrorLogHtml

	' Standard ASP cleanup of objects
	Set email = Nothing
	Set mailman = Nothing

%>



</BODY>
</HTML>
