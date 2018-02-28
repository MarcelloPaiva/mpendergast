<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>Send Japanese HTML Mail from ASP</TITLE>
</HEAD>
<BODY>

<%
	' Use ChilkatUtil.MbRequest to get the HTTP request data.
	set mbRequest = Server.CreateObject("ChilkatUtil.MbRequest")

	' Call BinaryRead to load the request data into the mbRequest object.
	mbRequest.BinaryRead

	' Create a mailman
	set mailman = Server.CreateObject("ChilkatWebMail.WebMailMan")

	' Unlock the component - use your trial or purchased unlock code.
	mailman.UnlockComponent "UnlockCode"

	' Tell the mailman where the SMTP server is located
	mailman.SmtpHost = "mail.earthlink.net"

	' Create an Email message
	set email = Server.CreateObject("ChilkatWebMail.WebEmail")

	' Enter the recipient's information
	email.AddMultipleTo email.QEncodeBytes(mbRequest.GetValue("Recipient"),"utf-8")

	' Enter the sender's information
	email.From = email.QEncodeBytes(mbRequest.GetValue("From"),"utf-8")

	' Enter the email subject
	email.Subject = email.QEncodeBytes(mbRequest.GetValue("Subject"),"utf-8")

	' Enter the email text
	email.SetMbHtmlBody "utf-8",mbRequest.GetValue("HtmlBody")

	' Convert the email to iso-2022-jp
	email.HtmlCharset = "iso-2022-jp"
	email.EncodeHeader "iso-2022-jp"
	
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
