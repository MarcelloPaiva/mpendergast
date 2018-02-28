<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>How to Read Email in ASP</TITLE>
</HEAD>
<BODY>

<%
	' Create a mailman
	set mailman = Server.CreateObject("ChilkatWebMail.WebMailMan")

	' Unlock the component - use your trial or purchased unlock code.
	mailman.UnlockComponent "unlock_code"

	mailman.MailHost = "mail.chilkatsoft.com"
	mailman.PopUsername = "myLogin"
	mailman.PopPassword = "myPassword"
	
	' Copy the email from the POP3 server without removing it.
	set bundle = mailman.GetAllHeaders(0)
	if (not (bundle is nothing)) then
		Response.write "<table border=1 cellpadding=5>"
		response.write "<tr>"
		response.write "<td><b>From</b></td>"
		response.write "<td><b>Subject</b></td>"
		response.write "<td><b>Display</b></td>"
		response.write "</tr>"
		
		' Loop over each email in the bundle.
		n = bundle.MessageCount
		for i = 0 to n-1
			set email = bundle.GetEmail(i)
			response.write "<tr>"
			response.write "<td>"+Server.HTMLEncode(email.From)+"</td>"
			response.write "<td>"+Server.HTMLEncode(email.Subject)+"</td>"
			response.write "<td><form method=""post"" action=""DisplayEmail.asp"" target=""emailDetail"">"
  			response.write "<input type=""submit"" name=""Submit"" value=""Display"">"
			response.write "<input type=""hidden"" name=""Uidl"" value=""" + email.Uidl +"""></form></td>"
			
			response.write "</tr>"
		next
		response.write "</table><br>"
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
