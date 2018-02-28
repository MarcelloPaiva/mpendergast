<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

	' Update our internal reference to be the first company record.
    xml.FirstChild2
    
    ' Find the next XML node where there is a child node having tag "name"
    ' with a value that matches the wildcarded pattern "Microsoft*"
    set microsoft = xml.FindNextRecord("name", "Microsoft*")
    If Not (microsoft Is Nothing) Then
        microsoft.AddAttribute "symbol", "MSFT"
    End If

	' Update our internal reference to be the XML document root.
	xml.GetRoot2
	
	response.write xml.GetXml()
%>


