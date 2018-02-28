<!--#include file="../globalFunctions/noCache.asp" -->
<%
	session.Contents.Remove("USR_" & session.SessionID)
	session.Contents.Remove("usr_name")
	session.Contents.Remove("usr_id")
	session.Contents.Remove("usr_arrAccess")
	
	session.Abandon()
	response.Redirect("default.asp")
%>