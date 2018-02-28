<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 13
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmState
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listState.asp")
	end if

	dim objState, id, abreviation, intName
	if validanumero(getQuery("id")) then
		set objState = new clsState
		objState.setID(getQuery("id"))
		objState.fndState()
		id = objState.getID
		abreviation = objState.getAbreviation
		intName = objState.getName
		set objState = nothing
	end if

%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= request.ServerVariables("SCRIPT_NAME") %>">
        <input name="id" type="hidden" id="id" value="<%= id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">STATE</td>
          </tr>
          <tr> 
            <td width="82">Name:</td>
            <td width="401"><input name="name" type="text" id="name" size="40" maxlength="50" value="<%= intName %>"></td>
          </tr>
          <tr> 
            <td>Abreviation:</td>
            <td><input name="abreviation" type="text" id="abreviation" size="10" maxlength="4" value="<%= abreviation %>"></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td align="right"><a href="listState.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0">
            </td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
