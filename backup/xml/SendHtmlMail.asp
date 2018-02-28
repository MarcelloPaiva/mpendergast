<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>Mail HTML with Embedded Image Example</TITLE>
</HEAD>
<BODY>

<%
	' Create a mailman
	set mailman = Server.CreateObject("ChilkatWebMail.WebMailMan")

	' Unlock the component - use your trial or purchased unlock code.
	mailman.UnlockComponent "unlock_code"

	' Tell the mailman where the SMTP server is located
	mailman.SmtpHost = "mail.earthlink.net"

	' Create an Email message
	set email = Server.CreateObject("ChilkatWebMail.WebEmail")

	' Enter the recipient's information
	email.AddTo "John Smith", "jsmith@chilkatsoft.com"

	' Enter the sender's information
	email.FromName = "Joe Smith"
	email.FromAddress = "joe.smith@somedomain.com"

	' Enter the email subject
	email.Subject = "Here is an image in HTML"

	' Modify this to point to your GIF image file.
	imageContentID = email.AddRelatedContent("c:\inetpub\wwwroot\images\ckActiveX.gif")

	' Enter the email text
	email.SetHtmlBody ("<HTML><HEAD></HEAD><BODY>"& _
"<br>This is an example of embedding an image in HTML email.<BR><IMG SRC="& _
chr(34)&"cid:"&imageContentID&chr(34)& _
"><br>The content ID of the image looks like this: "& _
imageContentID&"<br>The HTML for embedding the image looks like this: &lt;img src="& _
chr(34)&"cid:"&imageContentID&chr(34)&"&gt;<br></BODY></HTML>")

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
