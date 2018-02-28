<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/clsConnection.asp" -->
<!--#include file="globalFunctions/functions.asp" -->
<%
Dim rsList
Dim rsList_numRows

Set rsList = Server.CreateObject("ADODB.Recordset")
rsList.ActiveConnection = MM_clsConnection_STRING
rsList.Source = "SELECT pro_name, pro_id, pro_img1, pro_desc, pro_price  FROM tb_property  WHERE pro_status = true AND (pro_number = 0 or pro_number = 1)  ORDER BY pro_id DESC"
rsList.CursorType = 0
rsList.CursorLocation = 2
rsList.LockType = 1
rsList.Open()

rsList_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

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
        <tr> 
          <td colspan="5" valign="top" bgcolor="#E3C66E"><!--#include file="includes/menuTop_buyers.asp" --></td> 
          <td valign="top" bgcolor="#FFFFFF"><img src="images/cnt_topShadow.gif" width="566" height="10"></td> 
          <td width="1" rowspan="2" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="10"></td> 
        </tr> 
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#E3C66E" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="178" valign="top" bgcolor="#E3C66E"> 
            <!--#include file="includes/menu_buyers.asp" -->
          </td> 
          <td width="1" valign="top" bgcolor="#DCB93F" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#B06021" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td valign="top" bgcolor="#FFFFFF"> <!-- CONTENT AREA STARTS --> 
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="images/spacer.gif" width="20" height="68"></td> 
                  <td width="400"><img src="images/pg_ttl__featuredProperties.gif" width="400" height="68"> 
                    <!--PAGE TITLE --> </td> 
                  <td width="125"><img src="images/cnt_topLogo.gif" width="125" height="68"></td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="images/spacer.gif" width="10" height="1"></td> 
                  <td colspan="2"><img src="images/spacer.gif" width="10" height="1"></td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td>&nbsp;</td> 
                  <td colspan="2"> <!--REPEAT REGION STARTS--> 
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0"> 
                      <tr> 
                        <td colspan="4" class="cntHdr01"></td> 
                      </tr> 
                      <tr> 
                        <td colspan="4" class="cntHdr01"><img src="images/spacer.gif" width="10" height="10"></td> 
                      </tr> 
                      <% 
While ((Repeat2__numRows <> 0) AND (NOT rsList.EOF)) 
%> 
                      <tr> 
                        <td width="160" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="160">
							<a href="propertyDetails.asp?id=<%=(rsList.Fields.Item("pro_id").Value)%>" class="photo01"><img src="imagesDB/<%=(rsList.Fields.Item("pro_img1").Value)%>" alt="<%=(rsList.Fields.Item("pro_name").Value)%>" width="250" height="190" border="0"></a><br>							
							<img src="images/prp_pht_viewDetails.gif" width="151" height="13">							</td>
                          </tr>
                        </table></td> 
                        <td valign="top">&nbsp;</td> 
                        <td valign="top">
						<table width="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="cntHdr01"><a href="propertyDetails.asp?id=<%=(rsList.Fields.Item("pro_id").Value)%>" class="cntHdr01"><%=(rsList.Fields.Item("pro_name").Value)%></a></td>
  </tr>
  <tr>
    <td valign="top" class="cntHdr02"><%= FormatCurrency((rsList.Fields.Item("pro_price").Value), -1, -2, -2, -2) %></td>
  </tr>
  <tr>
    <td valign="top" class="cntHdr02">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" class="cntTxt01"><% if len(rsList.Fields.Item("pro_desc").Value) > 500 then response.write changeImage( mid(rsList.Fields.Item("pro_desc").Value,1 ,500) ) & "..." else response.write changeImage( rsList.Fields.Item("pro_desc").Value ) %>      <br>
      ( <a href="propertyDetails.asp?id=<%=(rsList.Fields.Item("pro_id").Value)%>">View Details </a> )</td>
  </tr>
</table>						</td> 
                        <td class="cntTxt01">&nbsp;</td> 
                      </tr> 
                      <tr>
                        <td width="160" valign="top"></td> 
                        <td width="10" valign="top"><img src="images/spacer.gif" width="10" height="10"></td> 
                        <td valign="top" class="cntHdr01">&nbsp;</td>
                        <td width="10" class="cntTxt01"><img src="images/spacer.gif" width="20" height="10"></td> 
                      </tr> 
                      <tr> 
                        <td valign="top"><img src="images/spacer.gif" width="10" height="20"></td> 
                        <td valign="top">&nbsp;</td> 
                        <td valign="top">&nbsp;</td> 
                        <td valign="top">&nbsp;</td> 
                      </tr> 
                      <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsList.MoveNext()
Wend
%> 
                    </table> 
                    <!--REPEAT REGION ENDS--> </td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td>&nbsp;</td> 
                  <td colspan="2">&nbsp;</td> 
                </tr> 
              </table> 
            </div> 
            <!-- CONTENT AREA ENDS --> </td> 
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
          <td valign="top" bgcolor="#FFFFFF" class="ftr01"><img src="images/ftr_legal_01.gif" name="ftr_legal" width="720" height="13" border="0" usemap="#ftr_legal" id="ftr_legal"><map name="ftr_legal"><area shape="rect" coords="679,1,719,12" href="adm/default.asp" target="_blank" alt="Click to login">
</map></td> 
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
rsList.Close()
Set rsList = Nothing
%>
