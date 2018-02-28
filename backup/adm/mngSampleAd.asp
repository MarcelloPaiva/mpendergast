<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 19
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmSampleAd
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listSampleAd.asp")
	end if
	
	dim obj, id, intTitle, image, doc, desc, url, intStatus
	intStatus = true
	
	if getQuery("action") = "delete" and validanumero(getQuery("id")) then
		set obj = new clsSampleAd
		obj.setID(getQuery("id"))
		obj.delSampleAd
		set obj = nothing
		
		response.Redirect("listSampleAd.asp")
	end if

	if validaNumero(getQuery("id")) then
		set obj = new clsSampleAd
		obj.setId(getQUery("id"))
		obj.fndSampleAd()
		id = obj.getID
		intTitle = obj.getTitle
		image = obj.getImage
		doc = obj.getDocument
		desc = obj.getDesc
		url = obj.getURL
		intStatus = obj.getStatus
		set obj = nothing
	end if
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form action="<%= REQUEST.ServerVariables("SCRIPT_NAME") %>" method="post" enctype="multipart/form-data" name="form1">
        <input type="hidden" name="id" value="<%= id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">SAMPLE AD</td>
          </tr>
          <tr> 
            <td>Title:</td>
            <td><input name="title" type="text" id="title" size="40" maxlength="100" value="<%= intTitle %>"></td>
          </tr>
          <tr> 
            <td>Image:</td>
            <td><input name="image" type="file" id="image"> <% if image <> "" then %>
              [<a href="view_image.asp?id=<%= id %>&section=sample_ad&image=<%= image %>&url=<%= request.ServerVariables("SCRIPT_NAME") & "?" & request.QueryString %>">view 
              image</a>] 
              <% end if %> </td>
          </tr>
          <tr>
            <td>Document:</td>
            <td><input name="document" type="file" id="document"> 
              <% if doc <> "" then %>
              [<a href="../imagesDB/<%= doc %>" target="_blank">view document</a>] 
              <% end if %>
            </td>
          </tr>
          <tr> 
            <td>URL:</td>
            <td><input name="url" type="url" id="title" size="40" maxlength="100" value="<%= url %>"></td>
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
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listTestimonial.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"> 
            </td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
