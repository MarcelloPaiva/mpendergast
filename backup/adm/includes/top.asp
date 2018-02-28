<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="204" valign="bottom"><script>doDate();</script> </td>
    <td width="546" align="right"><img src="../images/hdr_logo_01.gif" width="546" height="59"></td>
  </tr>
  <tr> 
    <td height="26" colspan="2" background="../images/tnav_bg.gif"><table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="10" class="tnav01"><img src="../images/spacer.gif" width="10" height="26"></td>
          <td width="508" class="tnav01">Welcome <% if session("usr_name") = "" then %>User<% else %><%= session("usr_name") %><% end if %></td>
          <td width="222" align="right"><% if session("usr_name") <> "" then %><a href="changePassword.asp" class="tnav01">Change your password</a><% end if %></td>
          <td width="10" class="tnav01"><img src="../images/spacer.gif" width="10" height="26"></td>

        </tr>
      </table></td>
  </tr>
</table>
