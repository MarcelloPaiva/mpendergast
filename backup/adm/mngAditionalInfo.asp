<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 9
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmAditionalInfo
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listAditionalInfo.asp")
	end if
	
	dim obj, id, title, desc, url, intStatus
	
	if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsAditionalInfo
		obj.setID(getQuery("id"))
		obj.delAditionalInfo
		set obj = nothing
		
		response.Redirect("listAditionalInfo.asp")
	end if
	
	intStatus = true
	if validaNumero(getQuery("id")) then
		set obj = new clsAditionalInfo
		obj.setId(getQUery("id"))
		obj.fndAditionalInfo()
		id = obj.getID
		title = obj.getTitle
		desc = obj.getDesc
		url = obj.getURL
		intStatus = obj.getStatus
		set obj = nothing
	end if
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= REQUEST.ServerVariables("SCRIPT_NAME") %>">
	  <input type="hidden" name="id" value="<%= id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">ADITIONAL INFO</td>
          </tr>
          <tr> 
            <td>Title:</td>
            <td><input name="title" type="text" id="title" size="40" maxlength="100" value="<%= title %>"></td>
          </tr>
          <tr> 
            <td>URL:</td>
            <td><input name="url" type="text" id="url" size="40" maxlength="100" value="<%= url %>"></td>
          </tr>
          <tr> 
            <td>Status:</td>
            <td><select name="status" id="status">
<%= getComboStatus(intStatus) %>
              </select></td>
          </tr>
          <tr> 
            <td>Description:</td>
            <td><textarea name="desc" cols="50" rows="10" id="desc"><%= desc %></textarea></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td align="right"><a href="listAditionalInfo.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0">
            </td>
          </tr>
        </table>
      </form> 
      <!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
