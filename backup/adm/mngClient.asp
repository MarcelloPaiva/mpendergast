<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 12
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmClient
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listClient.asp")
	end if
	
	dim obj, objCity, objClient, id, cit_id, intName, email, address, zipCode, areaCode, phone, joinDate

	if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsClient
		obj.setID(getQuery("id"))
		obj.delClient
		set obj = nothing
		
		response.Redirect("listClient.asp")
	end if

	if validaNumero(getQuery("id")) then
		set objClient = new clsClient
		objClient.setID(getQuery("id"))
		objClient.fndClient()
		id = objClient.getID
		cit_id = objClient.getCityID
		intName = objClient.getName
		email = objClient.getEmail
		address = objClient.getAddress
		zipCode = objClient.getZipCode
		areaCode = objClient.getAreaCode
		phone = objClient.getPhone
		joinDate = objClient.getJoinDate
		set  objClient = nothing
	end if
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= request.ServerVariables("SCRIPT_NAME") %>">
        <input name="id" type="hidden" id="id" value="<%= id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">CLIENT</td>
          </tr>
          <tr> 
            <td width="82">Name:</td>
            <td width="401"><input name="name" type="text" id="name" size="40" maxlength="100" value="<%= intName %>"></td>
          </tr>
          <tr> 
            <td>E-mail:</td>
            <td><input name="email" type="text" id="email" size="40" maxlength="100" value="<%= email %>"> 
            </td>
          </tr>
          <tr> 
            <td>City</td>
            <td><select name="cit_id" id="cit_id">
                <option value="">Select</option>
                <%
	set objCity = new clsCity
	response.Write objCity.getComboCity(cit_id)
	set objCity = nothing
%>
              </select></td>
          </tr>
          <tr> 
            <td>Address:</td>
            <td><input name="address" type="text" id="address" size="40" maxlength="100" value="<%= address %>"></td>
          </tr>
          <tr> 
            <td>Zip Code:</td>
            <td><input name="zipCode" type="text" id="zipCode" size="10" maxlength="8" value="<%= zipCode %>">
              (only numbers)</td>
          </tr>
          <tr> 
            <td>Phone:</td>
            <td><input name="areaCode" type="text" id="areaCode" size="4" maxlength="3" value="<%= areaCode %>"> 
              <input name="phone" type="text" id="phone" size="10" maxlength="8" value="<%= phone %>">
              (only numbers)</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listClient.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"></td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
