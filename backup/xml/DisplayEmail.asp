<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<%
	' Create a mailman
	set mailman = Server.CreateObject("ChilkatWebMail.WebMailMan")

	' Unlock the component - use your trial or purchased unlock code.
	mailman.UnlockComponent "unlock_code"

	mailman.MailHost = "mail.chilkatsoft.com"
	mailman.PopUsername = "myLogin"
	mailman.PopPassword = "myPassword"

	set email = mailman.FetchEmail(Request.Form("Uidl"))
	
	if (not (email is nothing)) then
		
		' Preferably display the HTML version
		if (email.HasHtmlBody() = 1) then
			Response.write email.GetHtmlBody()
		elseif (email.HasPlainTextBody() = 1) then
			%>
			<html>
			<head>
			<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;CHARSET=utf-8">
			<title>ASP to Display Plain-Text Email</title>
			</head>
			<body bgcolor="#FFFFFF">
<pre>
<%
' Get the plain-text body as utf-8 character data and HTML encoded
' This allows any email in any language to be displayed correctly.
Response.BinaryWrite email.GetMbPlainTextBody("utf-8",1)
%>
</pre>
			</body>
			</html>
		<%
		else
			Response.write "This email has no body."
		end if
		
	end if
%>
