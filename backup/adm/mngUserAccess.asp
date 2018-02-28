<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 8
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmUserAccess
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("mngUserAccess.asp?usr_id=" & getQuery("usr_id"))
	end if

	dim obj, usr_id, acs_id, uName
	usr_id = getQuery("usr_id")
	acs_id = getQuery("acs_id")

	if getQuery("action") = "delete" and validanumero(usr_id) and validanumero(acs_id) then
		set obj = new clsUser
		obj.setID(usr_id)
		obj.delUserAccess(acs_id)
		set obj = nothing
		
		response.Redirect("mngUserAccess.asp?usr_id=" & usr_id)
	end if
	
	set obj = new clsUser
	obj.setID(usr_id)
	obj.fndUser()
	uName = obj.getName
	set obj = nothing
	
	if not validaNumero(usr_id) then response.Redirect("listUser.asp")
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form action="<%= request.ServerVariables("SCRIPT_NAME") & "?usr_id=" & getQuery("usr_id") %>" method="post" name="form1">
        <input name="usr_id" type="hidden" value="<%= usr_id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">USER ACCESS</td>
          </tr>
          <tr> 
            <td width="82">User:</td>
            <td width="401"><%= uName %></td>
          </tr>
          <tr> 
            <td>Access:</td>
            <td><select name="acs_id" id="acs_id">
                <option value="">Select</option>
                <%
	set obj = new clsUser
	response.Write obj.getComboAccess(acs_id)
	set obj = nothing
%>
              </select></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listUser.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"></td>
          </tr>
        </table>
      </form> 
<%
	dim sql, conn, rs
	
	set rs = server.CreateObject("adodb.recordset")
	set conn = new clsConnection
	
	sql = "select a.acs_id, a.acs_section from tb_access a" &_
		" inner join tb_usr_acs ua ON ua.acs_id = a.acs_id" &_
		" where ua.usr_id = " & usr_id & " order by a.acs_section"
	rs.open sql, conn.conn, 3
%>
      <table width="450" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr align="center"> 
          <td colspan="2" class="txt3">User Access (<%= rs.recordCount %>)</td>
        </tr>
        <tr> 
          <td height="4" colspan="2" bgcolor="#CCCCCC"></td>
        </tr>
        <% if rs.bof and rs.eof then %>
        <tr align="center"> 
          <td height="70" colspan="2" class="txt1">Not found !!!</td>
        </tr>
        <%
	else
	while not rs.eof
%>
        <tr class="txt2"> 
          <td width="391">&raquo; <%= rs("acs_section") %></td>
          <td width="42" align="center"><a href="<%= request.ServerVariables("SCRIPT_NAME") %>?action=delete&usr_id=<%= usr_id %>&acs_id=<%= rs("acs_id") %>"><img src="images/remove.gif" alt="Remove" width="13" height="15" border="0"></a></td>
        </tr>
        <tr> 
          <td height="1" colspan="2" bgcolor="#CCCCCC"></td>
        </tr>
        <%
	rs.movenext
	wend
	end if
	rs.close
	set conn = nothing
%>
        <tr> 
          <td height="4" colspan="2" bgcolor="#CCCCCC"></td>
        </tr>
      </table>

<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
