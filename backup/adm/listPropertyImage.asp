<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 16
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
<%
	dim sql, pro_id, pro_name, iPageCurrent, iPageSize, iPageCount, iRecordsShown, conn, rs, search

	pro_id = getQuery("pro_id")
	if not validaNumero(pro_id) then response.Redirect("listProperty.asp")
	
	set rs = server.CreateObject("adodb.recordset")
	set conn = new clsConnection
	
	sql = "select pro_name from tb_property where pro_id = " & pro_id
	rs.open sql, conn.conn
	if not (rs.bof and rs.eof) then
		pro_name = rs("pro_name")
	end if
	rs.close
	
	if ValidaNumero(getQuery("page")) then
		iPageCurrent = cint(getQuery("page"))
	else
		iPageCurrent = 1
	end if

	iPageSize = 20
	rs.cursorlocation = 3

	sql = "select pi_id, pi_title, pi_image, pi_status from tb_property_image where pro_id = " & pro_id & " order by pi_title"
	rs.open sql, conn.conn

	rs.PageSize = iPageSize
	rs.CacheSize = iPageSize

	iPageCount = rs.PageCount

	If iPageCurrent > iPageCount Then
		iPageCurrent = iPageCount
	end if

	If iPageCurrent < 1 Then
		iPageCurrent = 1
	end if

	iRecordsShown = 0
%>
      <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr> 
          <td width="409" class="admHdr01">Photo Gallery IMAGES</td>
        </tr>
      </table> 
      <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr align="right">
          <td colspan="5"><a href="mngPropertyImage.asp?pro_id=<%= pro_id %>"><img src="images/new.gif" width="13" height="15" border="0" align="absmiddle"> Add new </a></td>
        </tr>
        <tr>
          <td colspan="5" class="cntHdr01"><%= pro_name %></td>
        </tr>
        <tr> 
          <td height="4" colspan="5" bgcolor="#CCCCCC"></td>
        </tr>
        <tr> 
          <td width="37" class="photo01">Preview</td>
          <td width="294" class="photo01">Name </td>
          <td width="26" align="center" class="photo01">Edit</td>
          <td width="21" align="center" class="photo01">Remove</td>
          <td width="34" align="center" class="photo01">On/Off</td>
        </tr>
        <tr> 
          <td height="1" colspan="5" bgcolor="#CCCCCC"></td>
        </tr>			
        <% if rs.bof and rs.eof then %>
        <tr align="center"> 
          <td height="70" colspan="5" class="txt1">No images posted for this property. </td>
        </tr>
        <%
	else
	rs.AbsolutePage = iPageCurrent
	While iRecordsShown < iPageSize and not rs.eof
%>
        <tr class="txt2"> 
          <td width="37"><img src="../imagesDB/<%= rs("pi_image") %>" width="37" height="30"></td>
          <td width="294"><%= rs("pi_title") %></td>
          <td align="center"><a href="mngPropertyImage.asp?pro_id=<%= pro_id %>&id=<%= rs("pi_id") %>"><img src="images/edit.gif" alt="View / Edit" width="13" height="15" border="0"></a></td>
          <td align="center"><a href="mngPropertyImage.asp?action=delete&pro_id=<%= pro_id %>&id=<%= rs("pi_id") %>"><img src="images/remove.gif" alt="Remove" width="13" height="15" border="0"></a></td>
          <td align="center"><%= getOnOffButton( rs("pi_id"), rs("pi_status") , "property_image") %></td>
        </tr>
        <tr> 
          <td height="1" colspan="5" bgcolor="#CCCCCC"></td>
        </tr>
        <%
	iRecordsShown = iRecordsShown + 1
	rs.movenext
	wend
	end if
	rs.close
	set conn = nothing
%>
        <tr> 
          <td height="4" colspan="5" bgcolor="#CCCCCC"></td>
        </tr>
        <tr align="center" class="txt1"> 
          <td colspan="5">Page <%= iPageCurrent %> of <%= iPageCount %> </td>
        </tr>
        <tr align="center" class="txt1"> 
          <td colspan="5"> <% If iPageCurrent > 1 Then %> <a href="<%= request.ServerVariables("PATH_INFO") & "?page=" & iPageCurrent - 1 & "&search=" & server.URLEncode(search) %>">&laquo; 
            previous</a> <% end if %> <% If iPageCurrent < iPageCount then %> <a href="<%= request.ServerVariables("PATH_INFO") & "?page=" & iPageCurrent + 1 & "&search=" & server.URLEncode(search) %>">Next 
            &raquo;</a> <% end if %> </td>
        </tr>
      </table>
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
