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
	xmlData = xml.GetURL("http://www.xml-parser.com/zipped.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    b = xml.FirstChild2()
    While (b)
        ' Unzip the sub-tree.
        xml.DecompressSubtree

		Response.write xml.GetChildContent("name") + "<br>"
		Response.write xml.GetChildContent("address") + "<br><br>"

        b = xml.NextSibling2()
    Wend
	
%>


</BODY>
</HTML>
