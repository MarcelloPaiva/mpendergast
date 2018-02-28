<!--#include file="../globalFunctions/noCache.asp" -->
<% if not session("USR_" & session.SessionID) then response.Redirect("default.asp") %>
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	dim fso, iFile, iFilePath, desc
	iFilePath = server.MapPath("../txt/txtForm.txt")
	
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		desc = getForm("desc")
		
		set fso = server.CreateObject("Scripting.FileSystemObject")
		set iFile = fso.CreateTextFile(iFilePath,true)
		iFile.write(desc)
		iFile.close
		set fso = nothing
		
		response.Redirect(request.ServerVariables("SCRIPT_NAME"))
	end if
	
	set fso = server.CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(iFilePath) then
		set iFile = fso.OpenTextFile(iFilePath,1)
		desc = iFile.ReadAll
		iFile.close
	end if
	set fso = nothing
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= REQUEST.ServerVariables("SCRIPT_NAME") %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">form congratulation text</td>
          </tr>
          <tr>
            <td>Text:</td>
            <td><textarea name="desc" cols="50" rows="7" id="desc"><%= desc %></textarea></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td width="399" align="right"><input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"> 
            </td>
          </tr>
        </table>
      </form> 
      <!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
