<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/clsConnection.asp" -->
<!--#include file="globalFunctions/functions.asp" -->
<%
Dim rsDetails__MMColParam
rsDetails__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsDetails__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsDetails
Dim rsDetails_numRows

Set rsDetails = Server.CreateObject("ADODB.Recordset")
rsDetails.ActiveConnection = MM_clsConnection_STRING
rsDetails.Source = "SELECT pro_desc, pro_id, pro_img1, pro_img2, pro_name, pro_price, pro_txt1, pro_txt2 FROM tb_property WHERE pro_id = " + Replace(rsDetails__MMColParam, "'", "''") + ""
rsDetails.CursorType = 0
rsDetails.CursorLocation = 2
rsDetails.LockType = 1
rsDetails.Open()

rsDetails_numRows = 0
%>
<%
Dim rsFeatures__MMColParam
rsFeatures__MMColParam = "0"
If (Request.QueryString("id") <> "") Then 
  rsFeatures__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rsFeatures
Dim rsFeatures_numRows

Set rsFeatures = Server.CreateObject("ADODB.Recordset")
rsFeatures.ActiveConnection = MM_clsConnection_STRING
rsFeatures.Source = "SELECT f.pf_name, pf.pp_desc  FROM tb_pro_pf pf  INNER JOIN tb_feature f on f.pf_id = pf.pf_id  WHERE pf.pro_id = " + Replace(rsFeatures__MMColParam, "'", "''") + " and pf.pp_status = true  ORDER BY f.pf_id ASC"
rsFeatures.CursorType = 0
rsFeatures.CursorLocation = 2
rsFeatures.LockType = 1
rsFeatures.Open()

rsFeatures_numRows = 0
%>
<%
Dim rsAerialImage__MMColParam
rsAerialImage__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  rsAerialImage__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rsAerialImage__MMstatus
rsAerialImage__MMstatus = "true"
If (true <> "") Then 
  rsAerialImage__MMstatus = true
End If
%>
<%
Dim rsAerialImage__MMaerial
rsAerialImage__MMaerial = "true"
If (true <> "") Then 
  rsAerialImage__MMaerial = true
End If
%>
<%
Dim rsAerialImage
Dim rsAerialImage_numRows

Set rsAerialImage = Server.CreateObject("ADODB.Recordset")
rsAerialImage.ActiveConnection = MM_clsConnection_STRING
rsAerialImage.Source = "SELECT pi_image  FROM tb_property_image  WHERE pro_id = " + Replace(rsAerialImage__MMColParam, "'", "''") + " and pi_status = " + Replace(rsAerialImage__MMstatus, "'", "''") + " and pi_aerial = " + Replace(rsAerialImage__MMaerial, "'", "''") + ""
rsAerialImage.CursorType = 0
rsAerialImage.CursorLocation = 2
rsAerialImage.LockType = 1
rsAerialImage.Open()

rsAerialImage_numRows = 0
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rsFeatures_numRows = rsFeatures_numRows + Repeat2__numRows
%>
<%
'Dim Repeat2__numRows
'Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rsList_numRows = rsList_numRows + Repeat2__numRows
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Martha Pendergast Real Estate -</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2" src="js/date.js" type="text/javascript"></script>
<script language="JavaScript1.2" src="js/global.js" type="text/javascript"></script>
<link href="css/style_01.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {
	color: #990000;
	font-weight: bold;
}
-->
</style>
</head>
<body onLoad="setHeight();MM_preloadImages('images/ftr_terms_01.gif')">
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0"> 
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td> 
    <td> <!-- HEADER STARTS --> 
      <table width="750" border="0" cellspacing="0" cellpadding="0"> 
        <tr> 
          <td valign="bottom"><script>doDate();</script> </td> 
          <td width="546" align="right"><img src="images/hdr_logo_01.gif" width="546" height="59"></td> 
        </tr> 
      </table> 
      <!-- HEADER ENDS --> </td> 
  </tr> 
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td> 
    <td> <!-- TOPNAV STARTS --> 
      <table width="750" height="26"  border="0" cellpadding="0" cellspacing="0" background="images/tnav_bg.gif"> 
        <tr> 
          <td><!-- #BeginLibraryItem "/Library/global_nav_01.lbi" --><table width="750" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="10"><img src="images/spacer.gif" width="10" height="26"></td>
              <td width="60" align="center"><a href="default.asp" class="tnav01">Home</a></td>
              <td width="80" align="center"><a href="genSections.asp?sec=5" class="tnav01">About Us</a> </td>
              <td width="170" align="center"><a href="portfolio.asp" class="tnav01">Portfolio of Sold Properties</a> </td>
              <td width="150" align="center"><a href="propertyList.asp" class="tnav01">Featured Properties</a> </td>
              <td width="180" align="center"><a href="contact2.asp" class="tnav01">Schedule an Appointment</a> </td>
              <td align="center"><a href="searchMLS.asp" class="tnav01">Search MLS</a> </td>
            </tr>
        </table><!-- #EndLibraryItem --></td> 
        </tr> 
      </table> 
      <!-- TOPNAV ENDS --> </td> 
  </tr> 
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td> 
    <td><img src="images/tnav_bottom.gif" width="750" height="6"></td> 
  </tr> 
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td> 
    <td valign="top"><table width="720" border="0" cellspacing="0" cellpadding="0">
        <!--DWLayoutTable-->
        <tr> 
          <td height="18" colspan="5" valign="top" bgcolor="#E3C66E"><!--#include file="includes/menuTop_buyers.asp" --></td>
          <td valign="top" bgcolor="#FFFFFF"><img src="images/cnt_topShadow.gif" width="566" height="10"></td>
          <td width="1" rowspan="2" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="10"></td>
        </tr>
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="1" valign="top" bgcolor="#E3C66E" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="178" valign="top" bgcolor="#E3C66E"><!--#include file="includes/menu_buyers_sold.asp" --></td>
          <td width="1" valign="top" bgcolor="#DCB93F" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="1" valign="top" bgcolor="#B06021" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td valign="top" bgcolor="#FFFFFF"> 
            <!-- CONTENT AREA STARTS -->
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="images/spacer.gif" width="20" height="68"></td>
                  <td width="400"><img src="images/pg_ttl__propertyDetails.gif" width="400" height="68"> 
                    <!--PAGE TITLE -->
                  </td>
                  <td width="125"><img src="images/cnt_topLogo.gif" width="125" height="68"></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="images/spacer.gif" width="10" height="25"></td>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td>&nbsp;</td>
                  <td colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="cntHdr01"><%=(rsDetails.Fields.Item("pro_name").Value)%> </td>
                      </tr>
                      <tr> 
                        <td colspan="2" class="cntHdr02">PROPERTY SOLD AT <%= FormatCurrency((rsDetails.Fields.Item("pro_price").Value), -1, -2, -2, -2) %></td>
                      </tr>
                      <tr>
                        <td colspan="2" class="cntTxt01">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="2" class="cntTxt01"><% if not (rsAerialImage.bof and rsAerialImage.eof) then %> <% if rsAerialImage.Fields.Item("pi_image").Value <> "" then %> <img src="imagesDB/<%=(rsAerialImage.Fields.Item("pi_image").Value)%>" width="500" border="0"> 
                        <% end if %> <% end if %> </td>
                      </tr>
                      <tr>
                        <td valign="bottom" class="cntTxt01">&nbsp;</td>
                        <td class="cntTxt01">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td valign="bottom" class="cntTxt01"><%= changeImage(rsDetails.Fields.Item("pro_desc").Value) %></td>
                        <td class="cntTxt01">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="2" class="cntHdr01"><img src="images/spacer.gif" width="10" height="20"></td>
                      </tr>
                      <tr> 
                        <td valign="top" class="cntTxt01">                          <% if rsDetails.Fields.Item("pro_img2").Value <> "" then %>                          <img src="imagesDB/<%=(rsDetails.Fields.Item("pro_img1").Value)%>" border="0" align="left" hspace="10">                          <% Else %>                          <img src="images/spacer.gif" width="1" height="1">                          <% End If %>                          <%= changeImage(rsDetails.Fields.Item("pro_txt1").Value)%> </td>
                        <td width="10" class="cntTxt01"><img src="images/spacer.gif" width="20" height="10"></td>
                      </tr>
                    </table>
                    <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td colspan="2" class="cntTxt01">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td valign="top" class="cntTxt01"> <p> 
                            <% if rsDetails.Fields.Item("pro_img2").Value <> "" then %>
                            <img src="imagesDB/<%=(rsDetails.Fields.Item("pro_img2").Value)%>" border="0" align="right" hspace="10"> 
                            <% Else %>
                            <img src="images/spacer.gif" width="1" height="1"> 
                            <% End If %>
                            <%= changeImage(rsDetails.Fields.Item("pro_txt2").Value)%> </p></td>
                        <td width="20" class="cntTxt01"><img src="images/spacer.gif" width="20" height="10"></td>
                      </tr>
                      <tr> 
                        <td valign="top"><img src="images/spacer.gif" width="10" height="30"></td>
                        <td valign="top">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td valign="top">
						<% if (NOT rsFeatures.EOF) then %>
								<table width="100%"  border="0" cellspacing="0" cellpadding="0">
								<tr> 
								<td colspan="2" class="cntFeaHdr01">Features</td>
								<td class="cntTxt01">&nbsp;</td>
								</tr>
									<% 
									cor = 1
									While ((Repeat2__numRows <> 0) AND (NOT rsFeatures.EOF)) %>
									<%	if cor = 1 then
									classe = "cntFea01"
									cor = 2
									else
									classe = "cntFea02"
									cor = 1
									end if
									%>
									<tr> 
									<td width="100" class="<%= classe %>"><%=(rsFeatures.Fields.Item("pf_name").Value)%></td>
									<td class="<%= classe %>"><%=(rsFeatures.Fields.Item("pp_desc").Value)%></td>
									<td width="10" class="cntTxt01"><img src="images/spacer.gif" width="20"></td>
									</tr>
									<% 
									Repeat2__index=Repeat2__index+1
									Repeat2__numRows=Repeat2__numRows-1
									rsFeatures.MoveNext()
									Wend
									%>
								<tr> 
								<td class="cntFeaFtr01">&nbsp;</td>
								<td class="cntFeaFtr01">&nbsp;</td>
								<td class="cntTxt01">&nbsp;</td>
								</tr>
								<tr> 
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td valign="top">&nbsp;</td>
								</tr>
								</table>                            
                            <% Else %>
                            <img src="images/spacer.gif" width="1" height="1"> 
                            <% End If %>						
						</td>
                        <td valign="top">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td valign="top" class="cntTxt01">Additional photos are available in the <a href="propertyPhotos.asp?id=<%=(rsDetails.Fields.Item("pro_id").Value)%>">Photo Gallery</a>.<br></td>
                        <td valign="top">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td valign="top">&nbsp;</td>
                        <td valign="top">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td>&nbsp;</td>
                  <td colspan="2">&nbsp;</td>
                </tr>
              </table>
            </div>
            <!-- CONTENT AREA ENDS -->
          </td>
        </tr>
      </table></td> 
  </tr> 
  <tr> 
    <td>&nbsp;</td> 
    <td valign="top"><table width="750" border="0" cellspacing="0" cellpadding="0"> 
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td valign="top" bgcolor="#FFFFFF" class="nws01"><table width="100%"  border="0" cellspacing="0" cellpadding="0"> 
              <tr> 
                <td height="25"><img src="images/nws_ttl_newsflash.gif" width="178" height="25"></td> 
                <td width="1"><img src="images/spacer.gif" width="1" height="1"></td> 
                <td><a href="testimonials.asp" onMouseOver="MM_swapImage('nws_ttl_testimonials','','images/nws_ttl_testimonials_02.gif',0)" onMouseOut="MM_swapImgRestore()"><img src="images/nws_ttl_testimonials.gif" name="nws_ttl_testimonials" width="208" height="25" border="0" id="nws_ttl_testimonials"></a></td> 
              </tr> 
              <tr> 
                <td width="40%" valign="top" class="nws02"><!--#include file="includes/news.asp" --></td> 
                <td width="1"><img src="images/nws_divider.gif" width="2" height="100"></td> 
                <td width="60%"><!--#include file="includes/testimonial.asp" --></td> 
              </tr> 
              <tr> 
                <td>&nbsp;</td> 
                <td><img src="images/spacer.gif" width="1" height="10"></td> 
                <td>&nbsp;</td> 
              </tr> 
            </table></td> 
          <td width="1" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="1"></td> 
        </tr> 
        <tr> 
          <td valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="15"></td> 
          <td valign="top" bgcolor="#FFFFFF" class="ftr01"><img src="images/ftr_legal_01.gif" name="ftr_legal" width="720" height="13" border="0" usemap="#ftr_legal" id="ftr_legal"></td> 
          <td valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="1"></td> 
        </tr> 
      </table></td> 
  </tr> 
</table> 
<br> 
<br> 
</body>
</html>
<%
rsDetails.Close()
Set rsDetails = Nothing
%>
<%
rsFeatures.Close()
Set rsFeatures = Nothing
%>
<%
rsAerialImage.Close()
Set rsAerialImage = Nothing
%>
