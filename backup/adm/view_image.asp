<!--#include file="../globalFunctions/noCache.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	dim id, page, table, imgN, image
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim obj, columnID, column, sql
		id = getForm("id")
		table = getForm("section")
		imgN = getForm("imgN") 'Para property: img1 e img2
		page = getForm("url")
		image = getForm("image")
		
		select case table
			case "sec_text"
				columnID = "st_id"
				column = "st_image"
			case "testimonial"
				columnID = "tes_id"
				column = "tes_image"
			case "property"
				columnID = "pro_id"
				if imgN = "2" then
					column = "pro_img2"
				else
					column = "pro_img1"
				end if
			case "sample_ad"
				columnID = "sa_id"
				column = "sa_image"
			case else
				response.Redirect("listNews.asp")
		end select
		
		set obj = new clsDelUploadedFile
		obj.delFile(server.MapPath("../imagesDB") & "\" & image)
		set obj = nothing
	
		sql = "update tb_" & trim(table) & " set " & column & " = null where " & columnID & " = " & id
		'response.Write(sql)
		'response.End()
		
		set obj = new clsConnection
		obj.conn.execute(sql)
		set obj = nothing
		
		response.Redirect(page)
	end if

	if not validaNumero(getQuery("id")) then response.Redirect("listNews.asp")

	id = getQuery("id")
	table = getQuery("section")
	imgN = getQuery("imgN") 'Para property: img1 e img2
	page = getQuery("url")
	image = getQuery("image")
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form name="form1" method="post" action="<%= request.ServerVariables("SCRIPT_NAME") %>">
	  <input type="hidden" name="id" value="<%= id %>">
	  <input type="hidden" name="section" value="<%= table %>">
	  <input type="hidden" name="imgN" value="<%= imgN %>">
	  <input type="hidden" name="url" value="<%= page %>">
	  <input type="hidden" name="image" value="<%= image %>">
        <table width="450" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td class="admHdr01">VIEW IMAGE</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td align="center"><img src="../imagesDB/<%= image %>"></td>
          </tr>
          <tr>
            <td align="center"><a href="javascript:document.forms[0].submit();">delete 
              this image</a></td>
          </tr>
        </table>
      </form> 
      <!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->