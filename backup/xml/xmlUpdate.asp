<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    ' Let's change "Adobe Systems Incorporated" to 
    ' "Adobe Systems Inc."

    ' Find the first node in the XML document where the tag is "name"
    ' and the content begins with "Adobe"
    set foundNode = xml.SearchForContent(Nothing, "name", "Adobe*")
    If (not (foundNode is Nothing)) Then
        foundNode.Content = "Adobe Systems Inc."
    End If

	response.write xml.GetXml()
%>


