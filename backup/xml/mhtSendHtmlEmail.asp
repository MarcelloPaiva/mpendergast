<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>Mail HTML with Embedded Image Example</TITLE>
</HEAD>
<BODY>

<%
	' Create a mailman
	set mht = Server.CreateObject("ChilkatMht.ChilkatMht")
	set mailman = Server.CreateObject("ChilkatWebMail.WebMailMan")

	' Unlock the component - use your trial or purchased unlock code.
	mailman.UnlockComponent "WebMailUnlockCode"
	mht.UnlockComponent "MhtUnlockCode"
	
	' Tell the mailman where the SMTP server is located
	mailman.SmtpHost = "mail.earthlink.net"

	' Create an Email message
	set email = mht.GetWebEmail("http://www.europe.reebok.com/EU/seeng/Home/default.htm")

	' Enter the recipient's information
	email.AddTo "John Smith", "jsmith@chilkatsoft.com"

	' Enter the sender's information
	email.FromName = "Joe Smith"
	email.FromAddress = "joe.smith@somedomain.com"

	' Enter the email subject
	email.Subject = "Reebok"

	' Send Email from ASP
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
