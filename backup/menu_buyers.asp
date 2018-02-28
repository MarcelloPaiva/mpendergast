<!--#include file="../classes/clsConnection.asp" -->
<!--#include file="menu_property.asp" -->
<table width="178" border="0" cellpadding="0" cellspacing="0" style="height: 100%;">
  <tr> 
    <td bgcolor="#FFFFFF"><img src="images/lnav_top_buyers_01.gif" name="lnav_top" width="183" height="21" border="0" usemap="#BUYERS_MAP" id="lnav_top"></td>
  </tr>
<%
	dim connMenu, rsMenu, sqlM
	set connMenu = new clsConnection
	set rsMenu = server.CreateObject("adodb.recordset")
	
	sqlM = "select pro_id, pro_name from tb_property where pro_status = true and pro_sellDate is null"
	rsMenu.open sqlM, connMenu.conn,3 
	if not (rsMenu.bof and rsMenu.eof) then
%>
  <tr> 
    <td><img src="images/lnav_ttl_featuredProperties.gif" width="179" height="35"></td>
  </tr>
<% while not rsMenu.eof %>
  <tr> 
    <td class="lnav01"><a href="?id=<%= rsMenu("pro_id") %>"><%= rsMenu("pro_name") %></a></td>
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
    <td class="lnav02"><a href="javascript:void(0);">Getting Ready to Buy</a> 
    </td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">Additional Information</a> 
    </td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">Industry News</a> </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr style="height: 100%;"> 
    <td><img id="glu" src="images/spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td><img src="images/lnav_ttl_quickMenu.gif" width="179" height="35"></td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">Home</a></td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">About Us</a> </td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">Portfolio</a></td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">Schedule an Appointment </a></td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">Featured Properties </a></td>
  </tr>
  <tr> 
    <td class="lnav02"><a href="javascript:void(0);">Search MLS</a> </td>
  </tr>
  <tr> 
    <td class="lnav02">&nbsp;</td>
  </tr>
</table>
<map name="BUYERS_MAP" id="BUYERS_MAP">
	<area shape="rect" coords="93,2,181,20" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_buyers_02.gif',0)" onMouseOut="MM_swapImgRestore()">
</map>