<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="includes/menu_property.asp" -->
<!--#include file="Connections/clsConnection.asp" -->
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
Dim rsPhoto__MMColParam
rsPhoto__MMColParam = "14"
If (Request.QueryString("id") <> "") Then 
  rsPhoto__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rsPhoto__MMstatus
rsPhoto__MMstatus = "true"
If (Request("MM_EmptyValue") <> "") Then 
  rsPhoto__MMstatus = Request("MM_EmptyValue")
End If
%>
<%
Dim rsPhoto
Dim rsPhoto_numRows

Set rsPhoto = Server.CreateObject("ADODB.Recordset")
rsPhoto.ActiveConnection = MM_clsConnection_STRING
rsPhoto.Source = "SELECT *  FROM tb_property_image  WHERE pro_id = " + Replace(rsPhoto__MMColParam, "'", "''") + " and pi_status = " + Replace(rsPhoto__MMstatus, "'", "''") + " order by pi_title"
rsPhoto.CursorType = 0
rsPhoto.CursorLocation = 2
rsPhoto.LockType = 1
rsPhoto.Open()

rsPhoto_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsProperties_numRows = rsProperties_numRows + Repeat1__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = -1
Repeat3__index = 0
rsPhoto_numRows = rsPhoto_numRows + Repeat3__numRows
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
          <td colspan="5" valign="top" bgcolor="#FFFFFF"><img src="images/lnav_top_buyers_01.gif" name="lnav_top" width="183" height="21" border="0" usemap="#Map" id="lnav_top"></td> 
          <td valign="top" bgcolor="#FFFFFF"><img src="images/cnt_topShadow.gif" width="566" height="10"></td> 
          <td width="1" rowspan="2" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="10"></td> 
        </tr> 
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#E3C66E"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="178" valign="top" bgcolor="#E3C66E">
            <!--#include file="includes/menu_buyers.asp" -->
          </td> 
          <td width="1" valign="top" bgcolor="#DCB93F"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#B06021"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td valign="top" bgcolor="#FFFFFF"> <!-- CONTENT AREA STARTS --> 
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="images/spacer.gif" width="20" height="68"></td> 
                  <td width="400"><img src="images/pg_ttl__photoGallery.gif" width="400" height="68"> 
                    <!--PAGE TITLE --> </td> 
                  <td width="125"><img src="images/cnt_topLogo.gif" width="125" height="68"></td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="images/spacer.gif" width="10" height="25"></td> 
                  <td colspan="2">&nbsp;</td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="images/spacer.gif" width="10" height="450"></td> 
                  <td colspan="2" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0"> 
                      <tr> 
                        <td class="cntHdr01"><%=(rsDetails.Fields.Item("pro_name").Value)%> </td> 
                      </tr>
                      <tr>
                        <td class="cntHdr02"><%= FormatCurrency((rsDetails.Fields.Item("pro_price").Value), -1, -2, -2, -2) %></td>
                      </tr>
                      <tr>
                        <td class="cntHdr01">&nbsp;</td>
                      </tr> 
                      <tr> 
                        <td class="cntTxt01"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="100%" align="center" valign="top" nowrap><% if NOT rsPhoto.EOF then %><img src="imagesDB/<%=(rsPhoto.Fields.Item("pi_image").Value)%>" name="viewImage" width="420" border="2" id="viewImage" alt="<%=(rsPhoto.Fields.Item("pi_title").Value)%>">
                              <% end if %></td>
                            <td align="right" valign="top"><table width="82" border="0" align="center" cellpadding="2" cellspacing="0">
                              
                                <% 
While ((Repeat3__numRows <> 0) AND (NOT rsPhoto.EOF)) 
%>
                                <tr>
								<td align="center" valign="top" class="photo01"><a href="javascript:;" onMouseOver="MM_swapImage('viewImage','','imagesDB/<%=(rsPhoto.Fields.Item("pi_image").Value)%>',0)"><img src="imagesDB/<%=(rsPhoto.Fields.Item("pi_image").Value)%>" alt="<%=(rsPhoto.Fields.Item("pi_title").Value)%>" name="thumb" width="40" height="35" border="0" id="thumb"></a></td>
                                </tr>
                                <tr>
                                    <td align="center" valign="top" class="photo01" style="position:relative; top:-8px"><%=(rsPhoto.Fields.Item("pi_title").Value)%></td>
                                  </tr>
		<% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  rsPhoto.MoveNext()
Wend
%>
                            </table></td>
                            <td align="right" valign="top"><img src="images/spacer.gif" width="10" height="25"></td>
                          </tr>
                          <tr>
                            <td align="center" valign="top" nowrap>&nbsp;</td>
                            <td align="right" valign="top">&nbsp;</td>
                            <td align="right" valign="top">&nbsp;</td>
                          </tr>
                          <tr>
                            <td align="center" valign="top" nowrap><a href="propertyDetails.asp?id=<%=(rsDetails.Fields.Item("pro_id").Value)%>">Back to <%=(rsDetails.Fields.Item("pro_name").Value)%> detail page</a> </td>
                            <td align="right" valign="top">&nbsp;</td>
                            <td align="right" valign="top">&nbsp;</td>
                          </tr>
                        </table></td> 
                      </tr> 
                    </table> 
                  </td>
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
          <td valign="top" bgcolor="#FFFFFF" class="ftr01"><img src="images/ftr_legal_01.gif" name="ftr_legal" width="720" height="13" border="0" usemap="#ftr_legal" id="ftr_legal"></td> 
          <td valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="1"></td> 
        </tr> 
      </table></td> 
  </tr> 
</table> 
<map name="Map"> 
  <area shape="rect" coords="93,2,181,20" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_buyers_02.gif',0)" onMouseOut="MM_swapImgRestore()"> 
</map> 
<br> 
<br> 
</body>
</html>
<%
rsDetails.Close()
Set rsDetails = Nothing
%>
<%
rsPhoto.Close()
Set rsPhoto = Nothing
%>

