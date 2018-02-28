<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 18
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmNews
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listNews.asp")
	end if
	
	dim obj, id, title, desc, reference, intStatus, endDate

	if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsNews
		obj.setID(getQuery("id"))
		obj.delNews
		set obj = nothing
		
		response.Redirect("listNews.asp")
	end if
	
	intStatus = true
	if validaNumero(getQuery("id")) then
		set obj = new clsNews
		obj.setId(getQuery("id"))
		obj.fndNews()
		id = obj.getID
		title = obj.getTitle
		desc = obj.getDesc
		reference = obj.getReference
		intStatus = obj.getStatus
		endDate = obj.getEndDate
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
            <td colspan="2" class="admHdr01">NEWS</td>
          </tr>
          <tr> 
            <td>Title:</td>
            <td><input name="title" type="text" id="title" size="40" maxlength="100" value="<%= title %>"></td>
          </tr>
          <tr> 
            <td>Reference:</td>
            <td><input name="reference" type="text" id="reference" size="40" maxlength="100" value="<%= reference %>"></td>
          </tr>
          <tr> 
            <td>Status:</td>
            <td><select name="status" id="status" class="<%= intStatus %>">
                <%= getComboStatus(intStatus) %> </select></td>
          </tr>
          <tr> 
            <td>Description:</td>
            <td><textarea name="desc" cols="50" rows="10" id="desc"><%= desc %></textarea></td>
          </tr>
          <tr>
            <td>Expiration Date:</td>
            <td> 
              <select name="Month" id="Month">
                <option value="">Month</option>
                <%= getComboMonth(endDate) %> </select> 
              <select name="Day" id="Day">
                <option value="">Day</option>
                <%= getComboDay(endDate) %> </select> 
              <input name="Year" type="text" id="Year" size="4" maxlength="5" <% if isDate(endDate) then %>value="<%= year(endDate) %>"<% end if %>></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listNews.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"> 
            </td>
          </tr>
        </table>
      </form> 
      <!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
