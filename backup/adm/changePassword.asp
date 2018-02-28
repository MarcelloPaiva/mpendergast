<!--#include file="../globalFunctions/noCache.asp" -->
<%
	if not session("USR_" &session.SessionID) then response.Redirect("logout.asp")

	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmUserChangePassword
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("logout.asp")
	end if
%>
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= request.ServerVariables("SCRIPT_NAME") %>">
        <table width="350" border="0" align="center" cellpadding="3" cellspacing="2">
          <tr> 
            <td colspan="2" class="admHdr01">Change your password</td>
          </tr>
          <tr>
            <td colspan="2">After change the passord, you have to login again.</td>
          </tr>
          <tr> 
            <td width="67">Atual Password:</td>
            <td width="265"><input name="pass1" type="password" id="pass1" size="15" maxlength="10"></td>
          </tr>
          <tr> 
            <td>New Password:</td>
            <td><input name="pass2" type="password" id="pass2" size="15" maxlength="10"></td>
          </tr>
          <tr> 
            <td>Confirm new password:</td>
            <td><input name="pass3" type="password" id="pass3" size="15" maxlength="10"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"></td>
          </tr>
        </table>
      </form> 
      <!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
