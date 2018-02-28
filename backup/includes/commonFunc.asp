<%
function lastNews()
	Dim result, strHref, strText, strTarget
	Dim oXML, oXMLError, ReturnValue, x
	result = ""
	
	Set oXML = Server.CreateObject("MSXML2.DOMDocument")
	oXML.setProperty "ServerHTTPRequest", true
	oXML.async = False
	ReturnValue = oXML.Load("http://www.inman.com/inf/marthapendergastrealestate/")

	Set objLst = oXML.getElementsByTagName("InmanStory")
	noOfHeadlines = objLst.length

	For i = 0 To (noOfHeadlines-1)
		Set objHdl = oXML.documentElement.childNodes(i)
		nwsTitle = trim(objHdl.childNodes(0).text)
		nwsDate = trim(objHdl.childNodes(2).text)
		nwsSubtitle = trim(objHdl.childNodes(1).text)
		nwsAuthor = trim(objHdl.childNodes(3).text)
		nwsArticle = trim(objHdl.childNodes(4).text)
		nwsSource = trim(objHdl.childNodes(5).text)

		strHref = "javascript:PopNews(" & i + 1 & ")"
		strText = ClearForJavascript("<span style='position:relative; left:-10px; line-height:30px; font-size:9px;'>") & nwsDate & "</span><br>" & "<b>" & ClearForJavascript(nwsTitle) & "</b><br>" & ClearForJavascript(ReadMore(nwsSubtitle & "...", 80, "... [read more]"))
		strTarget = "_self"
		result = result & "['" & strHref & "','" & strText & "','" & strTarget & "']," & Chr(13)

		'Response.Write( "<h1>" & (i+1) &")" & nwsTitle & "</h1><br>")
		'Response.Write( nwsSubtitle & "<br>")
		'Response.Write( nwsDate & "<br>")  
		'Response.Write( nwsAuthor & "<br>")  
		'Response.Write( nwsSource & "<br>")  
		'Response.Write( nwsArticle & "<br>")
		Set objHdl = nothing
	Next

	If result <> "" Then
		result = "[" & Left(result, Len(result)-2) & "]"
	End If
	
	set objLst = nothing
	Set oXML = nothing

	lastNews = result
end function

Function ReadMore(content, limit, append)
  Dim result, contentaux
  contentaux = Trim(content)
  If limit > Len(contentaux) Then
     result = contentaux
  Else
     result = Mid(contentaux, 1, InstrRev(contentaux, " ", limit)-1) & append
  End If
  ReadMore = result
End Function


Function ClearForJavascript(content)
   Dim result
   result = Replace(content, "'", "\'")
   result = Replace(result, Chr(13), "\n")
   result = Replace(result, Chr(10), "")
   ClearForJavascript = result
End Function

'response.write lastNews
%>