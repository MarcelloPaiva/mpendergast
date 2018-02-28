<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<TITLE>Mail HTML with Embedded Image Example</TITLE>
</HEAD>
<BODY>

<%
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    ' Find all companies located in California
    ' Navigate to the first company record.
    xml.FirstChild2()

    ' Find the next XML node where there is a child node having tag "name"
    ' with a value that matches the wildcarded pattern "Microsoft*"
    Response.write "<h2>Companies with Headquarters in California</h2>"

    While Not (xml Is Nothing)
        ' FindNextRecord *will* return the current record if it
        ' matches the criteria. 
        set xml = xml.FindNextRecord("state", "CA")

        If Not (xml Is Nothing) Then
            ' Add the company name to the listbox.
            Response.write(xml.GetChildContent("name") + "<br>") 

            ' Advance past this record.
            set xml = xml.NextSibling()
        End If

    Wend
	
%>


</BODY>
</HTML>
