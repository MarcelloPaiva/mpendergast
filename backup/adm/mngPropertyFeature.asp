<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 15
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmPropertyFeature
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("mngPropertyFeature.asp?pro_id=" & getQuery("pro_id"))
	end if

	dim obj, id, pro_id, pName, desc, intStatus
	id = getQuery("id")
	pro_id = getQuery("pro_id")
	intStatus = true
	
	if getQuery("action") = "changeStatus" and validanumero(id) and validanumero(pro_id) then
	
		if getQuery("status") = "1" then
			intStatus = 1
		else
			intStatus = 0
		end if
		
		set conn = new clsConnection
		conn.conn.execute("update tb_pro_pf set pp_status = " & intStatus & " where pro_id = " & pro_id & " and pf_id = " & id)
		set conn = nothing
		
		response.Redirect(request.ServerVariables("SCRIPT_NAME") & "?pro_id=" & pro_id)
	end if
	
	if getQuery("action") = "delete" and validanumero(id) and validanumero(pro_id) then
		set obj = new clsPropertyFeature
		obj.setPropertyID(pro_id)
		obj.setFeatureID(id)
		obj.delPropertyFeature
		set obj = nothing
		
		response.Redirect("mngPropertyFeature.asp?pro_id=" & pro_id )
	end if

	if not validaNumero(pro_id) then response.Redirect("listProperty.asp")
	
	if validanumero(id) then
		set obj = new clsPropertyFeature
		obj.setFeatureID(id)
		obj.setPropertyID(pro_id)
		obj.fndPropertyFeature
		desc = obj.getDesc
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
      <form action="<%= request.ServerVariables("SCRIPT_NAME") & "?pro_id=" & getQuery("pro_id") %>" method="post" name="form1">
        <input name="pro_id" type="hidden" id="pro_id" value="<%= pro_id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">PROPERTY'S FEATURE</td>
          </tr>
          <tr> 
            <td width="82">Property:</td>
            <td width="401" class="cntHdr01"><%= pName %></td>
          </tr>
          <tr> 
            <td>Title:</td>
            <td><select name="pf_id" id="pf_id">
                <option value="">Select</option>
<%
	set obj = new clsFeature
	response.Write obj.getComboFeature(id)
	set obj = nothing
%>
              </select></td>
          </tr>
		<tr> 
            <td>Status:</td>
            <td><select name="status" id="status">
				<%= getComboStatus(intStatus) %>
              </select></td>
          </tr>
          <tr> 
            <td>Description:</td>
            <td><textarea name="desc" cols="50" rows="5" id="desc"><%= desc %></textarea></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"></td>
          </tr>
        </table>
      </form> 
<%
	dim sql, conn, rs
	
	set rs = server.CreateObject("adodb.recordset")
	set conn = new clsConnection
	
	sql = "select f.pf_name, pp.pf_id, pp.pp_status from tb_pro_pf pp" &_
		" inner join tb_feature f ON f.pf_id = pp.pf_id" &_
		" where pp.pro_id = " & pro_id & " order by f.pf_name"
	rs.open sql, conn.conn, 3
%>
      <table width="450" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr align="center"> 
          <td colspan="4" class="txt3">Property Features (<%= rs.recordCount %>)</td>
        </tr>
        <tr> 
          <td height="1" colspan="4" bgcolor="#CCCCCC"><img src="../images/spacer.gif"></td>
        </tr>
        <% if rs.bof and rs.eof then %>
        <tr align="center"> 
          <td height="70" colspan="4" class="txt1">Not found !!!</td>
        </tr>
        <%
	else
	while not rs.eof
%>
        <tr class="txt2"> 
          <td width="316">&raquo; <a href="<%= request.ServerVariables("SCRIPT_NAME") %>?pro_id=<%= pro_id %>&pf_id=<%= rs("pf_id") %>"><%= rs("pf_name") %></a></td>
          <td width="31" align="center"><a href="<%= request.ServerVariables("SCRIPT_NAME") %>?pro_id=<%= pro_id %>&id=<%= rs("pf_id") %>"><img src="images/edit.gif" alt="View / Edit" width="13" height="15" border="0"></a></td>
          <td width="36" align="center"><a href="<%= request.ServerVariables("SCRIPT_NAME") %>?action=delete&pro_id=<%= pro_id %>&id=<%= rs("pf_id") %>"><img src="images/remove.gif" alt="Remove" width="13" height="15" border="0"></a></td>
          <td width="32" align="center">
		  	<% if rs("pp_status") then %>
				<a href="?action=changeStatus&status=0&pro_id=<%= pro_id %>&id=<%= rs("pf_id") %>"><img src="images/on.gif" border="0"></a>
			<% else %>
				<a href="?action=changeStatus&status=1&pro_id=<%= pro_id %>&id=<%= rs("pf_id") %>"><img src="images/off.gif" border="0"></a>
			<% end if %>
		  </td>
        </tr>
        <tr> 
          <td height="1" colspan="4" bgcolor="#CCCCCC"></td>
        </tr>
        <%
	rs.movenext
	wend
	end if
	rs.close
	set conn = nothing
%>
        <tr> 
          <td height="1" colspan="4" bgcolor="#CCCCCC"><img src="../images/spacer.gif"></td>
        </tr>
      </table>

<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
