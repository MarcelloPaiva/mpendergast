<!--#include file="../globalFunctions/noCache.asp" -->
<% 
	dim section_code
	select case getQuery("section")
		case "2"
			section_code = 2
		case "3"
			section_code = 3
		case "4"
			section_code = 4
		case "5"
			section_code = 5
		case "6"
			section_code = 6
		case else
			section_code = 1
	end select
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmSecText
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listSecText.asp?section=" & getQuery("section"))
	end if
	
	dim obj, id, section, sectionName, title, image, desc, footer, intStatus, url

	section = getQuery("section")
	if not ValidaNumero(section) then section = 1

	if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsSecText
		obj.setID(getQuery("id"))
		obj.delSecText
		set obj = nothing
		
		response.Redirect("listSecText.asp?section=" & section)
	end if

	intStatus = true
	if validaNumero(getQuery("id")) then
		set obj = new clsSecText
		obj.setId(getQUery("id"))
		obj.fndSecText()
		id = obj.getID
		section = obj.getSectionID
		title = obj.getTitle
		image = obj.getImage
		desc = obj.getDesc
		footer = obj.getFooter
		intStatus = obj.getStatus
		url = obj.getURL
		set obj = nothing
	end if
	
	set obj = new clsSecText
	sectionName = obj.fndSectionName(section)
	set obj = nothing
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form action="<%= REQUEST.ServerVariables("SCRIPT_NAME") & "?section=" & section %>" method="post" enctype="multipart/form-data" name="form1">
        <input type="hidden" name="id" value="<%= id %>">
		<input type="hidden" name="section" value="<%= section %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01"><%= ucase(sectionName) %></td>
          </tr>
          <tr> 
            <td>Title:</td>
            <td><input name="title" type="text" id="title" size="40" maxlength="100" value="<%= title %>"></td>
          </tr>
          <% if cint(section) = 3 then %>
          <tr> 
            <td>Footer:</td>
            <td><input name="footer" type="text" id="footer" size="40" maxlength="255" value="<%= footer %>"></td>
          </tr>
          <% end if %>
          <tr> 
            <td>Image:</td>
            <td><input type="file" name="file"> <% if image <> "" then %>
              [<a href="view_image.asp?id=<%= id %>&section=sec_text&image=<%= image %>&url=<%= request.ServerVariables("SCRIPT_NAME") & "?" & request.QueryString %>">view 
              image</a>]
              <% end if %></td>
          </tr>
          <tr> 
            <td>Status:</td>
            <td><select name="status" id="status">
                <%= getComboStatus(intStatus) %> </select></td>
          </tr>
          <tr> 
            <td>Description:</td>
            <td><textarea name="desc" cols="50" rows="10" id="desc"><%= desc %></textarea></td>
          </tr>
          <% if cint(section) = 2 then %>
          <tr>
            <td>URL:</td>
            <td><input name="url" type="text" id="url" size="40" maxlength="100" value="<%= url %>"></td>
          </tr>
		<% end if %>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listSecText.asp?section=<%= getQuery("section") %>"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"> 
            </td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
