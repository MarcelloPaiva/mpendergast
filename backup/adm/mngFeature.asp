<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 11
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmFeature
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listFeature.asp")
	end if

	dim objFeature, id, intName
	
		if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsFeature
		obj.setID(getQuery("id"))
		obj.delFeature
		set obj = nothing
		
		response.Redirect("listFeature.asp")
	end if
	
	
	if validanumero(getQuery("id")) then
		set objFeature = new clsFeature
		objFeature.setID(getQuery("id"))
		objFeature.fndFeature()
		id = objFeature.getID
		intName = objFeature.getName
		set objFeature = nothing
	end if

%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= request.ServerVariables("SCRIPT_NAME") %>">
        <input name="id" type="hidden" id="id" value="<%= id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">FEATURE</td>
          </tr>
          <tr> 
            <td width="82">Name:</td>
            <td width="401"><input name="name" type="text" id="name" size="40" maxlength="50" value="<%= intName %>"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listFeature.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"> 
            </td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
