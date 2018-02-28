<!--#include file="../globalFunctions/noCache.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<% 
	dim section_code
	section_code = 15
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
<%
	dim sql, iPageCurrent, iPageSize, iPageCount, iRecordsShown, conn, rs, search
	
	set rs = server.CreateObject("adodb.recordset")
	set conn = new clsConnection
	
	if ValidaNumero(getQuery("page")) then
		iPageCurrent = cint(getQuery("page"))
	else
		iPageCurrent = 1
	end if

	iPageSize = 20
	rs.cursorlocation = 3

	sql = "select pro_id, pro_name, pro_status from tb_property where pro_selldate is null order by pro_name"
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
      <table width="550" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr> 
          <td width="409" class="admHdr01">PROPERTIES</td>
        </tr>
      </table> 
      <table width="450" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr> 
          <td height="4" colspan="6" bgcolor="#CCCCCC"><img src="../images/spacer.gif" height="1"></td>
        </tr>
        <tr> 
          <td width="23" class="photo01">Edit</td>
          <td width="215" class="photo01">Property Name </td>
          <td width="50" align="center" class="photo01">Features</td>
          <td width="38" align="center" class="photo01">Images</td>		  
          <td width="41" align="center" class="photo01">Remove</td>
          <td width="38" align="center" class="photo01">On/Off</td>
        </tr>
        <tr> 
          <td height="1" colspan="6" bgcolor="#CCCCCC"></td>
        </tr>			
        <% if rs.bof and rs.eof then %>
        <tr align="center"> 
          <td height="70" colspan="6" class="txt1">No properties found.</td>
        </tr>
        <%
	else
	rs.AbsolutePage = iPageCurrent
	While iRecordsShown < iPageSize and not rs.eof
%>
        <tr class="txt2"> 
          <td><a href="mngProperty.asp?id=<%= rs("pro_id") %>"><img src="images/edit.gif" alt="EDIT" width="13" height="15" border="0"></a></td>
          <td width="215"><a href="mngProperty.asp?id=<%= rs("pro_id") %>"><%= rs("pro_name") %></a></td>
          <td align="center"><a href="mngPropertyFeature.asp?pro_id=<%= rs("pro_id") %>"><img src="images/features.gif" alt="FEATURES" border="0"></a></td>
          <td align="center"><a href="listPropertyImage.asp?pro_id=<%= rs("pro_id") %>"><img src="images/images.gif" alt="PHOTOS" border="0"></a></td>
          <td align="center"><a href="mngProperty.asp?action=delete&id=<%= rs("pro_id") %>"><img src="images/remove.gif" alt="Remove" width="13" height="15" border="0"></a></td>
          <td align="center"><%= getOnOffButton( rs("pro_id"), rs("pro_status") , "property") %></td>
        </tr>
        <tr> 
          <td height="1" colspan="6" bgcolor="#CCCCCC"></td>
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
          <td height="1" colspan="6" bgcolor="#CCCCCC"><img src="../images/spacer.gif" height="1"></td>
        </tr>
        <tr align="center" class="txt1"> 
          <td colspan="6">Page <%= iPageCurrent %> of <%= iPageCount %> </td>
        </tr>
        <tr align="center" class="txt1"> 
          <td colspan="6"> <% If iPageCurrent > 1 Then %> <a href="<%= request.ServerVariables("PATH_INFO") & "?page=" & iPageCurrent - 1 & "&search=" & server.URLEncode(search) %>">&laquo; 
            previous</a> <% end if %> <% If iPageCurrent < iPageCount then %> <a href="<%= request.ServerVariables("PATH_INFO") & "?page=" & iPageCurrent + 1 & "&search=" & server.URLEncode(search) %>">Next 
            &raquo;</a> <% end if %> </td>
        </tr>
      </table>
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
