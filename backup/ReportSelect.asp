<%@language = vbscript%>

<%
	Session("PAGETITLE")="Tally Workflow Reports Menu"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><html>
<head><title><%=Session("PAGETITLE")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=unicode">
<meta NAME="GENERATOR" Content="MSHTML 6.00.2479.6">
</head>
<body>
<LINK REL="stylesheet" HREF="inc/Workflow.css">
<!--#INCLUDE FILE="inc/vbscript_general.asp" -->
<!--#INCLUDE FILE="inc/header.asp" -->
<!--#INCLUDE FILE="inc/toolbar.asp" -->

<p align="center">
<table id="TABLE1" cellSpacing="1" cellPadding="1" border="1">

	<tr>
		<td colspan=1><font size="5">Select From Available Reports:</font></td>
	</tr>
	<tr>
		<td colspan=1><font size="5">&nbsp</font></td>
	</tr>

<%
	dim rs
	dim cmd
	
	cmd = "usp_GetReports"
	set rs = oConn.Execute(cmd)
	if NOT rs.eof then
		do while not rs.eof
%>

	<tr>
		<td colSpan="2" align=center><font size="5">
		<% If ISNULL(rs.fields("SelectionPage")) Then %>
			<A href="ReportView.asp?RptID=<%=rs.fields("ID")%>"><%=rs.fields("Description")%></A></font>
		<% else %>
			<A href="<%=rs.fields("SelectionPage")%>?Tgt=ReportView.asp?RptID=<%=rs.fields("ID")%>"><%=rs.fields("Description")%></A></font>
		<% end if %>
		</td>
	</tr>
<%
		rs.movenext
		loop
	end if
%>
</table></p>

</body>

<!--#INCLUDE FILE="inc/footer.asp" -->

</html>
