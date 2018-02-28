<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

	' Send this XML to another ASP page for processing.
	' In this case, it simply echos the XML back.
	set xml2 = xml.HttpPost("http://127.0.0.1/Examples/xmlEcho.asp")

	response.write xml2.GetXml()
%>


