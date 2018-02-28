<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>AES Interop Example</TITLE>
</HEAD>
<BODY bgcolor="#FFFFFF">
<%
	' Create the Chilkat object.
	set aes = Server.CreateObject("AesInterop.AesInterop")

	' The password must be the same for decryption.
	aes.PassPhrase = "MySecretPassword"

	' Use 256-bit encryption
	aes.KeyLength = 256

	' Data to encrypt
	myData = "Chilkat Software, Inc."
%>

<h2><font color="#330099">1. AES Encrypt a String: &quot;Chilkat Software, Inc.&quot;</font></h2>
<h3>AES Encrypted and then Base64 Encoded:</h3>
<%
	strBase64 = aes.EncodeToBase64(aes.Encrypt(myData)) 
	Response.write strBase64
%>
<br><br>
<h3>AES Encrypted and then Hex-Encoded:</h3>
<%
	strHex = aes.EncodeToHex(aes.Encrypt(myData)) 
	Response.write strHex
%>
<br><br>
<h3>AES Encrypted and then Converted to Quoted-Printable:</h3>
<%
	strQP = aes.EncodeToQP(aes.Encrypt(myData)) 
	Response.write strQP
%>
<br><br>
<h2><font color="#330099">2. Decrypt Each String</font></h2>
<h3>Decode from Base64 and AES Decrypt:</h3>
<%
	str = aes.Decrypt(aes.DecodeFromBase64(strBase64))
	Response.write str
%>
<h3>Hex-Decode and then AES Decrypt:</h3>
<%
	str = aes.Decrypt(aes.DecodeFromHex(strHex))
	Response.write str
%>
<h3>QP-Decode and then AES Decrypt:</h3>
<%
	str = aes.Decrypt(aes.DecodeFromQP(strQP))
	Response.write str
%>

</BODY>
</HTML>
