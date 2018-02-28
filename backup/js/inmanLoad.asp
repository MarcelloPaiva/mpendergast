<%
Dim oXML, oXMLError, ReturnValue, x
Set oXML = Server.CreateObject("MSXML2.DOMDocument")
oXML.setProperty "ServerHTTPRequest", true
oXML.async = False
ReturnValue = oXML.Load("http://www.inman.com/inf/marthapendergastrealestate/")

Set objLst = oXML.getElementsByTagName("InmanStory")
noOfHeadlines = objLst.length

' Dim InmanStory, headline, subtitle, date, byline, bodycontent, copyright
'	InmanStory = oXML.documentElement.childNodes(i).text
'	headline = oXML.documentElement.childNodes(i).childNodes(0).text
'	subtitle = oXML.documentElement.childNodes(i).childNodes(1).text
'	date = oXML.documentElement.childNodes(i).childNodes(2).text
'	byline = oXML.documentElement.childNodes(i).childNodes(3).text
'	bodycontent = oXML.documentElement.childNodes(i).childNodes(4).text
'	copyright = oXML.documentElement.childNodes(i).childNodes(5).text		
'	Set oXML = Nothing


Function lastNews()
     Dim result, strHref, strText, strTarget
     result = ""
	 
	 i = 0
	While not( i > (noOfHeadlines-1))
				 
		strHref = "javascript:void(0);"
		'strText = "<table><tr><td class=newsHP width=12><img src=images/bullet_black.gif border=0 width=20 height=20 align=bottom></td><td width=200 valign=middle class=newsHP><b>" & rsLastNews("nws_Headline") & "</b></td></tr><tr><td></td><td>" & ReadMore(rsLastNews("nws_Article"), 500, "... (click for more)") & "</td></tr><tr><td colspan=3 align=center><img src=images/news_bar.gif border=0 width=150 height=20></td></tr></table>"
		Set objHdl = oXML.documentElement.childNodes(4)
		nwsTitle = objHdl.childNodes(0).text
		nwsDate = objHdl.childNodes(2).text
		nwsSubtitle = objHdl.childNodes(1).text
		nwsAuthor = objHdl.childNodes(3).text
		nwsArticle = objHdl.childNodes(5).text
		nwsSource = objHdl.childNodes(4).text
		
		strText = "nwsTitle"
		'strText = "<img src=images/bullet_black_news.gif border=0 width=20 height=120 align=left>" &_
		'		 "<b>" & nwsTitle & "</b><br><br>" & ClearForJavascript(ReadMore( nwsArticle , 180, "... [read more]"))
		'		 "<b>" & nwsTitle & "</b><br><br>" & ClearForJavascript(ReadMore( nwsArticle , 180, "... [read more]"))
		strTarget = "_blank"
		result = "['','" & strText & "','" & strTarget & "']," & Chr(13)

     i = i+1
     Wend
     If result <> "" Then
        result = "[" & Left(result, Len(result)-2) & "]"
     End If
     lastNews = result
	 
	 Set oXML = Nothing

End Function

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


Function ReadAllTextFile(textFile)
  Const ForReading = 1
  Dim result, fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(textFile, ForReading)
  result = f.ReadAll
  f.Close
  ReadAllTextFile = result
End Function
%>



	