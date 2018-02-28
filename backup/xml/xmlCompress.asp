<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

	' Navigate to the first company record.
    b = xml.FirstChild2()
    While (b)
        ' Save the company name as an attribute.
        xml.AddAttribute "name", xml.GetChildContent("name")

        ' Zip compress the sub-tree.
        xml.CompressSubtree

        b = xml.NextSibling2()
    Wend

	' Navigate back to the XML Document root.
    xml.GetRoot2

	response.write xml.GetXml()
%>


