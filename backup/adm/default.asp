<!--#include file="../globalFunctions/noCache.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim login, pass, erro, obj, return
		
		login = getForm("login")
		pass = getForm("pass")
		
		if len(login) < 4 then erro = erro & "Invalid user name.<br>"
		if len(pass) < 4 then erro = erro & "Invalid password.<br>"
		
		if erro <> "" then
			fcErro(erro)
		end if
		
		set obj = new clsUser
		return = obj.chkUser(login, pass)
		if return then
			session("usr_name") = obj.getName
			session("usr_id") = obj.getID
			session("usr_arrAccess") = obj.arrAccess
			session("USR_" & session.SessionID) = true
		else
			set obj = nothing
			fcErro("Invalid user or passowrd")
		end if
		set obj = nothing
		
		response.redirect("listProperty.asp")
	end if
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300" bgcolor="#FFFFFF"><form name="form1" method="post" action="<%= request.ServerVariables("SCRIPT_NAME") %>">
        <table width="350" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">Restricted Area</td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td width="156" align="right">login:</td>
            <td width="294"><input name="login" type="text" id="login" size="20" maxlength="20"></td>
          </tr>
          <tr> 
            <td align="right">password:</td>
            <td><input name="pass" type="password" id="pass" size="20" maxlength="10"></td>
          </tr>
          <tr>
            <td align="right">&nbsp;</td>
            <td align="right">&nbsp;</td>
          </tr>
          <tr>
            <td align="right">&nbsp;</td>
            <td>
              <input name="Submit" type="submit" class="Button" value="   Login   ">
            </td>
          </tr>
        </table>
      </form>
      
    </td>
  </tr>
</table>
<!--#include file="includes/bottom.asp" -->
