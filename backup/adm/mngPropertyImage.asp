<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 16
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmPropertyImage
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listPropertyImage.asp?pro_id=" & getQuery("pro_id"))
	end if

	dim obj, id, pro_id, pName, title, desc, image, main, aerial, intStatus
	id = getQuery("id")
	pro_id = getQuery("pro_id")

	if getQuery("action") = "delete" and validanumero(id) then
		set obj = new clsPropertyImage
		obj.setID(id)
		obj.delPropertyImage
		set obj = nothing
		
		response.Redirect("listPropertyImage.asp?pro_id=" & pro_id)
	end if

	if not validaNumero(id) and not validaNumero(pro_id) then response.Redirect("mngProperty.asp")
	
	if validanumero(id) then
		set obj = new clsPropertyImage
		obj.setID(id)
		obj.fndPropertyImage
		pro_id = obj.getPropertyID
		title = obj.getTitle
		desc = obj.getDesc
		image = obj.getImage
		main = obj.getMain
		aerial = obj.getAerial
		intStatus = obj.getStatus
		set obj = nothing
	end if
	
	if validaNumero(pro_id) then
		set obj = new clsProperty
		obj.setID(pro_id)
		obj.fndProperty
		pName = obj.getName
		set obj = nothing
	end if
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form action="<%= request.ServerVariables("SCRIPT_NAME") & "?pro_id=" & pro_id %>" method="post" enctype="multipart/form-data" name="form1">
        <input name="id" type="hidden" id="id2" value="<%= id %>">
        <input name="pro_id" type="hidden" id="id3" value="<%= pro_id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">Photo Gallery IMAGE</td>
          </tr>
          <tr>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" class="cntHdr01"><%= pName %></td>
          </tr>
          <tr>
            <td colspan="2">Photos should be <span style="font-weight: bold">420 pixels</span> wide max. </td>
          </tr>
          <tr>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td width="82">Title:</td>
            <td width="401"><input name="title" type="text" id="title" size="40" maxlength="100" value="<%= title %>"></td>
          </tr>
          <tr> 
            <td>Image:</td>
            <td><input name="file" type="file" id="file"> 
              <span class="photo01">(420 pixels wide)</span> </td>
          </tr>
          <tr> 
            <td>Aerial:</td>
            <td><input name="aerial" type="checkbox" id="aerial" value="1" <% if aerial then %> checked<% end if %>> 
              <span class="photo01">(Check only if photo is an aerial view) </span></td>
          </tr>
          <tr> 
            <td>Active:</td>
            <td><select name="status" id="status">
              <%= getComboStatus(intStatus) %>
            </select></td>
          </tr>
<% if image <> "" then %>
          <tr>
            <td colspan="2">Image Preview: </td>
          </tr>
          <tr align="center"> 
            <td colspan="2"><img src="../imagesDB/<%= image %>" width="233" height="167" alt="<%= image %>"><br>
              <span class="photo01">Photo gallery images should be 420 pixels wide.</span></td>
          </tr>
          <tr align="center">
            <td colspan="2" class="admHdr01">&nbsp;</td>
          </tr>
<% end if %>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listPropertyImage.asp?pro_id=<%= getQuery("pro_id") %>"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"></td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
