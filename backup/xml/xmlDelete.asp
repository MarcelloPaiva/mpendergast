<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    xml.FirstChild2
    set microsoft = xml.FindNextRecord("name", "Microsoft*")

    ' The Microsoft record is now root itself.
    microsoft.RemoveFromTree

    ' Save the XML with the Microsoft record removed.
    xml.GetRoot2

	response.write xml.GetXml()
%>


