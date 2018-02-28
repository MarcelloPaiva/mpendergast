<!--#include file="../globalFunctions/noCache.asp" -->
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<%
	dim erro
	erro = trim(request.QueryString("erro"))
%>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="300"><table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
        <tr> 
          <td class="admHdr01">Attention:</td>
        </tr>
<% if erro = "" then %>
        <tr> 
          <td><font color="#FF0000">An error has occurred.</font></td>
        </tr>
<% else %>
        <tr> 
          <td><font color="#FF0000"><%= erro %></font></td>
        </tr>
<% end if %>
        <tr> 
          <td><input name="Button" type="button" class="Button" onClick="history.back();" value="   Back   "></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<!--#include file="includes/bottom.asp" -->
