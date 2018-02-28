<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    ' Insert a new company
    set companyNode = xml.NewChild("company", "")
    set nameNode = companyNode.NewChild("name", "Apple Computer, Inc.")

    ' Add an attribute to this node.
    nameNode.AddAttribute "Symbol", "AAPL"

    companyNode.NewChild2 "address", "1 Infinite Loop"
    companyNode.NewChild2 "city", "Cupertino"
    companyNode.NewChild2 "state", "CA"
    companyNode.NewChild2 "zip", "95014"
    companyNode.NewChild2 "website", "http:www.apple.com"
    companyNode.NewChild2 "phone", "408-996-1010"

	response.write xml.GetXml()
%>


