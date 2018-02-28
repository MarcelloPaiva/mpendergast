<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

	' Returns the XML page as a Variant
	xmlData = xml.GetURL("http://www.xml-parser.com/companies.xml","","","")
	
	' Make sure you use LoadXmlV (for Variant) instead of LoadXml (for String)
	xml.LoadXmlV xmlData

    ' Quickly locate the Chilkat record.
 	set ckRec = xml.SearchForContent(Nothing, "name", "Chilkat*")
    If (not (ckRec is Nothing)) Then

        ' Move up to the record level.
        ckRec.GetParent2

        set gifNode = ckRec.NewChild("gif_image", "")

        ' The data can be optionally Zip compressed and/or AES encrypted.
        gifData = xml.GetURL("http://www.chilkatsoft.com/images/dude.gif","","","")
        
	    zipFlag = False
	    aesEncryptFlag = False
	    password = ""
        gifNode.SetBinaryContent zipFlag, aesEncryptFlag, password, gifData

    End If

	response.write xml.GetXml()
%>


