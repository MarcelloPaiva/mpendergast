<%
Dim oXML, oXMLError, ReturnValue, x
Set oXML = Server.CreateObject("MSXML2.DOMDocument")
oXML.setProperty "ServerHTTPRequest", true
oXML.async = False

ReturnValue = oXML.Load("http://www.inman.com/inf/marthapendergastrealestate/")
Dim nws_headline, nws_subtitle, nws_date, nws_author, nws_article, nws_copyright
	nws_headline = oXML.documentElement.childNodes(0).childNodes(0).text
	nws_subtitle = oXML.documentElement.childNodes(0).childNodes(1).text
	nws_date = oXML.documentElement.childNodes(0).childNodes(2).text
	nws_author = oXML.documentElement.childNodes(0).childNodes(3).text
	nws_article = oXML.documentElement.childNodes(0).childNodes(4).text
	nws_copyright = oXML.documentElement.childNodes(0).childNodes(5).text		
	Set oXML = Nothing
%>