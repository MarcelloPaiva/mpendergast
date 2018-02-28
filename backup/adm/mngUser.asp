<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 8
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmUser
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listUser.asp")
	end if
	
	dim obj, id, intName, login, intStatus
	intStatus = true
	
	if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsUser
		obj.setID(getQuery("id"))
		obj.delUser
		set obj = nothing
		
		response.Redirect("listUser.asp")
	end if
	
	if validaNumero(getQuery("id")) then
		set obj = new clsUser
		obj.setId(getQUery("id"))
		obj.fndUser()
		id = obj.getID
		intName = obj.getName
		login = obj.getLogin
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
            <td colspan="2" class="admHdr01">USER</td>
          </tr>
          <tr> 
            <td>Name:</td>
            <td width="399"> <input name="name" type="text" id="name" size="40" maxlength="100" value="<%= intName %>"> 
            </td>
          </tr>
          <tr> 
            <td>Login:</td>
            <td><input name="login" type="text" id="login" size="40" maxlength="20" value="<%= login %>"></td>
          </tr>
          <tr> 
            <td>Status:</td>
            <td><select name="status" id="status">
                <%= getComboStatus(intStatus) %> </select></td>
          </tr>
<% if validaNumero(id) then %>
          <tr align="center" bgcolor="efefef"> 
            <td colspan="2">Only fill the password fields if you want update it 
              !!! </td>
          </tr>
<% end if %>
          <tr> 
            <td>Password:</td>
            <td><input name="pass1" type="password" id="pass1" size="10" maxlength="8"></td>
          </tr>
          <tr> 
            <td>Confirm Password:</td>
            <td><input name="pass2" type="password" id="pass2" size="10" maxlength="8"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listUser.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"> 
            </td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
