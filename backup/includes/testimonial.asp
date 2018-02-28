<%
set conn = Server.CreateObject("ADODB.connection")
Set rsShowTestimonial = Server.CreateObject("ADODB.Recordset")
conn.open(MM_clsConnection_STRING)

rsShowTestimonial.open  "select tes_desc, tes_name, tes_image from tb_testimonial", conn,3
iRowCount = rsShowTestimonial.recordCount
if iRowCount <> 0 then
	randomize
	iRowCount = int(iRowCount * rnd())
	rsShowTestimonial.move iRowCount

	tesName = rsShowTestimonial("tes_name")
	tesDesc = rsShowTestimonial("tes_desc")
	tesImage = trim(rsShowTestimonial("tes_image"))
end if

rsShowTestimonial.close
conn.close
set rsShowTestimonial = nothing
set conn = nothing

if len(tesDesc) > 100 then tesDesc = mid(tesDesc, 1, 190) & "..."
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td valign="top" class="tst01"><em><%= tesDesc %></em> ( <a href="testimonials.asp">more</a> )</td>
    <td valign="top"><% if tesImage <> "" then %><img src="imagesDB/<%= tesImage %>" width="160" height="108"><% end if %></td>
    <td><img src="images/spacer.gif" width="10" height="10"></td>
  </tr>
</table>