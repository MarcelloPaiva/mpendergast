<!--#include file="../globalFunctions/noCache.asp" -->
<% 
	dim section_code
	section_code = 14
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
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

			sql = "select s.stt_abrev, c.cit_id, c.cit_name from tb_city c" &_
				" inner join tb_state s on s.stt_id = c.stt_id" &_
				" order by s.stt_abrev, c.cit_name"
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
          <td width="409" class="admHdr01">CITIES</td>
          <td width="141" align="right"><a href="mngCity.asp"><img src="images/new.gif" width="13" height="15" border="0" align="absmiddle"> 
            Add new </a></td>
        </tr>
      </table> 
      <table width="450" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr> 
          <td height="1" colspan="3" bgcolor="#CCCCCC"><img src="../images/spacer.gif"></td>
        </tr>
        <% if rs.bof and rs.eof then %>
        <tr align="center"> 
          <td height="70" colspan="3" class="txt1">Not found !!!</td>
        </tr>
        <%
	else
	rs.AbsolutePage = iPageCurrent
	While iRecordsShown < iPageSize and not rs.eof
%>
        <tr class="txt2"> 
          <td width="355">&raquo; <%= rs("stt_abrev") %> / <%= rs("cit_name") %></td>
          <td width="39" align="center"><a href="mngCity.asp?id=<%= rs("cit_id") %>"><img src="images/edit.gif" alt="View / Edit" width="13" height="15" border="0"></a></td>
          <td width="32" align="center"><a href="mngCity.asp?action=delete&id=<%= rs("cit_id") %>"><img src="images/remove.gif" alt="Remove" width="13" height="15" border="0"></a></td>
        </tr>
        <tr> 
          <td height="1" colspan="3" bgcolor="#CCCCCC"></td>
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
          <td height="1" colspan="3" bgcolor="#CCCCCC"><img src="../images/spacer.gif"></td>
        </tr>
        <tr align="center" class="txt1"> 
          <td colspan="3">Page <%= iPageCurrent %> of <%= iPageCount %> </td>
        </tr>
        <tr align="center" class="txt1"> 
          <td colspan="3"> <% If iPageCurrent > 1 Then %> <a href="<%= request.ServerVariables("PATH_INFO") & "?page=" & iPageCurrent - 1 & "&search=" & server.URLEncode(search) %>">&laquo; 
            previous</a> <% end if %> <% If iPageCurrent < iPageCount then %> <a href="<%= request.ServerVariables("PATH_INFO") & "?page=" & iPageCurrent + 1 & "&search=" & server.URLEncode(search) %>">Next 
            &raquo;</a> <% end if %> </td>
        </tr>
      </table>
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
