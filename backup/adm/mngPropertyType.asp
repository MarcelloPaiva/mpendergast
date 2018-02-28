<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 17
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmPropertyType
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listPropertyType.asp")
	end if

	dim obj, id, intName
	
		if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsPropertyType
		obj.setID(getQuery("id"))
		obj.delPropertyType
		set obj = nothing
		
		response.Redirect("listPropertyType.asp")
	end if
	
	if validaNumero(getQuery("id")) then
		set obj = new clsPropertyType
		obj.setId(getQuery("id"))
		obj.fndPropertyType()
		id = obj.getID
		intName = obj.getName
		set obj = nothing
	end if
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= request.ServerVariables("SCRIPT_NAME") %>">
        <input name="id" type="hidden" id="id" value="<%= id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">PROPERTY TYPE</td>
          </tr>
          <tr> 
            <td width="82">Name:</td>
            <td width="401"><input name="name" type="text" id="name" size="40" maxlength="50" value="<%= intName %>"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listPropertyType.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"> 
            </td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
