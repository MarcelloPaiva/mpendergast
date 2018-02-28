<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>Unzip XML Records On-The-Fly</TITLE>
</HEAD>
<BODY>

<%
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.inman.com/inf/marthapendergastrealestate","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    set company = xml.FirstChild()

    While Not (company Is Nothing)
        Response.write company.GetChildContent("name") + "<br>"
        set company = company.NextSibling()
    Wend
	
%>


</BODY>
</HTML>
