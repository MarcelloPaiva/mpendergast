<!--#include file="../classes/clsConnection.asp" -->
<%
	function menuProperty(id)
		dim html

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyDetails.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0" 
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "detail") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Property Details</a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyPhotos.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "photos") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Photo Gallery </a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyVtour.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "vtour") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Virtual Tour </a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyAddInfo.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "addinfo") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Additional Information </a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='contact.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "contact") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Schedule an Appointment </a></td> "
		html = html & "</tr> "

		menuProperty = html
	end function
%>
<table width="178" border="0" cellpadding="0" cellspacing="0" style="height: 100%;">
  <%
	dim connMenu, rsMenu, sqlM
	set connMenu = new clsConnection
	set rsMenu = server.CreateObject("adodb.recordset")
	
	sqlM = "select pro_id, pro_name from tb_property where pro_status = true and pro_sellDate is null ORDER BY pro_name ASC"
	rsMenu.open sqlM, connMenu.conn,3 
	if not (rsMenu.bof and rsMenu.eof) then
%>
  <tr> 
    <td><a href="propertyList.asp"><img src="images/lnav_ttl_featuredProperties.gif" width="179" height="35" border="0"></a></td>
  </tr>
  <% while not rsMenu.eof %>
  <tr> 
    <td class="lnav01"><a href="propertyDetails.asp?id=<%= rsMenu("pro_id") %>"><%= rsMenu("pro_name") %></a></td>
  </tr>
  <% if trim(request.QueryString("id")) = cstr(rsMenu("pro_id")) then response.write menuProperty(trim(request.QueryString("id"))) %>
  <%
	rsMenu.movenext
	wend
	end if
	rsMenu.close
	set connMenu = nothing
%>
  <tr> 
    <td class="lnavBorder">&nbsp; </td>
  </tr>
  <tr> 
    <td><img src="images/lnav_ttl_prepareYourself.gif" width="179" height="35"></td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="genSections.asp?sec=1">Getting Ready to Buy</a> 
    </td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="propertyAddInfo.asp">Additional Information</a> 
    </td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:PopNews(1)">Industry News</a> </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
<!--#include file="menu_quick.asp" -->