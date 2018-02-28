<!--#include file="../classes/clsConnection.asp" -->
<%
	function menuSellers()
		dim html

'- know the value... ( o mesmo de buyers)
'- choosing an agent
'- useful links
'- request ...
'- sample ads
'- contact

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='genSections.asp?sec=3'><img src='images/lnav_sub_bullet_0" 
		if trim(request.QueryString("sec")) = "3" then
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Know your property value</a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='genSections.asp?sec=4'><img src='images/lnav_sub_bullet_0" 
		if trim(request.QueryString("sec")) = "4" then
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Choosing an Agent</a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='sellersusefulLinks.asp'><img src='images/lnav_sub_bullet_0" 
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "useful") <> 0 then
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Useful Links</a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='contact3.asp'><img src='images/lnav_sub_bullet_0" 
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "request") <> 0 then
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Request a Market Analysis</a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='sampleAds.asp'><img src='images/lnav_sub_bullet_0" 
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "sample") <> 0 then
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Sample Ads</a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='contact2.asp'><img src='images/lnav_sub_bullet_0" 
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "contact") <> 0 then
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Contact Us</a></td> "
		html = html & "</tr> "

		menuSellers = html
	end function
%>
<table width="178" border="0" cellpadding="0" cellspacing="0" style="height: 100%;">
  <tr> 
    <td><img src="images/lnav_ttl_gettingReadyToSell.gif" width="179" height="35"></td>
  </tr>
  <%= menuSellers %> 
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
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
    <td><a href="propertyList.asp"><img src="images/lnav_ttl_featuredProperties.gif" width="179" height="35" border="0"></a></td>
  </tr>
  <% while not rsMenu.eof %>
  <tr> 
    <td class="lnav01"><a href="propertyDetails.asp?id=<%= rsMenu("pro_id") %>"><%= rsMenu("pro_name") %></a></td>
  </tr>
  <%
	rsMenu.movenext
	wend
	end if
	rsMenu.close
	set connMenu = nothing
%>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td class="lnavBorder">&nbsp;</td>
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
</table>
<!--#include file="menu_quick.asp" -->
