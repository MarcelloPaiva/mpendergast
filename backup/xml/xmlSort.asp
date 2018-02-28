<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    ' Sort the company records by name.
    ascending = True
    xml.SortRecordsByContent "name", ascending

    ' Sort the fields within each company record by the field name.
    b = xml.FirstChild2()
    While b
        xml.SortByTag ascending
        b = xml.NextSibling2()
    Wend

	' Navigate back to the root.
    xml.GetRoot2()

	response.write xml.GetXml()
%>


