<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<%
	id = trim(request.QueryString("id"))
	if id <> "" and isNumeric(id) then id = cint(id)
	
	Set oXML = Server.CreateObject("MSXML2.DOMDocument")
	oXML.setProperty "ServerHTTPRequest", true
	oXML.async = False
	ReturnValue = oXML.Load("http://www.inman.com/inf/marthapendergastrealestate/")

	Set objLst = oXML.getElementsByTagName("InmanStory")
	noOfHeadlines = objLst.length

	Set objHdl = oXML.documentElement.childNodes(id - 1)
	nwsTitle = trim(objHdl.childNodes(0).text)
	nwsDate = trim(objHdl.childNodes(2).text)
	nwsSubtitle = trim(objHdl.childNodes(1).text)
	nwsAuthor = trim(objHdl.childNodes(3).text)
	nwsArticle = trim(objHdl.childNodes(4).text)
	nwsSource = trim(objHdl.childNodes(5).text)
	Set objLst = nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Martha Pendergast NEWS</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/style_02.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--

.style1 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style3 {color: #0066CC}
img {
	border: 10px solid #FFFFFF;
	top: 16px;
	position: relative;
}
.divider {
	border: 0px solid #cccccc;
	top: 0px;
	position: relative;
}

-->
</style></head>
<body onLoad="window.focus();" class="newsPopBody">
<br>
<table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" class="newsPopBack"> 
  <tr>
    <td align="center"><img src="images/newsHeader.gif" width="491" height="57"></td>
  </tr>
  <tr> 
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="10">
      <tr>
        <td><span class="newsPopTitle"><br>
          <%= nwsTitle %></span><br>
		  <!-- <span class="newsPopDescr"><%= nwsAuthor %></span></td> -->
      </tr>
    </table></td> 
  </tr> 
  <tr> 
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="10">
        <tr>
          <td valign="top">
            <table width="200" border="0" align="left" cellpadding="3" cellspacing="0" bgcolor="#F6F6F6">
              <tr> 
                <td bgcolor="#EEEEEE"><div align="left" class="style1">More 
                    News</div></td>
                <td width="1" bgcolor="#FFFFFF"><img src="images/spacer.gif" width="1" height="1" align="left"></td>
              </tr>
<%
	For i = 0 To (noOfHeadlines-1)
		Set objHdl = oXML.documentElement.childNodes(i)
%>
              <tr> 
                <td><a href="?id=<%= i + 1 %>" class="newsPopHP"><%= trim(objHdl.childNodes(0).text) %></a></td>
                <td bgcolor="#FFFFFF" class="newsPopHP">&nbsp;</td>
              </tr>
              <tr>
                <td align="center"><img src="images/newsDivider.gif" width="150" height="1" class="divider"></td>
                <td bgcolor="#FFFFFF" class="newsPopHP">&nbsp;</td>
              </tr>
<%
		Set objHdl = nothing
	Next
	Set oXML = nothing
%>
            </table>
          <a href="javascript:void(0);print();" class="newsPopHP style3">CLICK HERE TO PRINT</a><br><br> 
          <div align="justify" class="newsPopDescr"><%= nwsDate %><%= nwsArticle + "</a>" %></div><br><br><%= nwsSource %></td>
        </tr>
      </table></td> 
  </tr> 
</table> 
  
</body>
</html>
