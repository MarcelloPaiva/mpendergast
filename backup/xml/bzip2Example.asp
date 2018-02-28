<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>ASP BZIP2 Compression Example</TITLE>
</HEAD>
<BODY bgcolor="#FFFFFF">
<p> 
<%
	' Create the Chilkat object
	set aes = Server.CreateObject("AesInterop.AesInterop")

	' Generate some data.
	myData = "ABCD 1234 wxyz<br>"
	for i = 0 to 8
		myData = myData + myData
	next
%> 
</p>
<h1>BZip2 Compressed and Base64 Encoded</h1>
<p> 
<%
	compressed = aes.EncodeToBase64(aes.BZip2(myData))
	Response.write compressed
%> 
</p>
<h1>Bzip2 Uncompressed</h1>
<p> 
<%
	uncompressed = aes.BUnzip2(aes.DecodeFromBase64(compressed))
	Response.write uncompressed
%> 
</p>
</BODY>
</HTML>
