<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/clsConnection.asp" -->
<%
Dim rsCAgent__MMColParam
rsCAgent__MMColParam = "true"
If (Request("MM_EmptyValue") <> "") Then 
  rsCAgent__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsCAgent
Dim rsCAgent_numRows

Set rsCAgent = Server.CreateObject("ADODB.Recordset")
rsCAgent.ActiveConnection = MM_clsConnection_STRING
rsCAgent.Source = "SELECT top 1 st.st_desc, st.st_image, st.st_title  FROM tb_sec_text st  inner join tb_section s ON s.sec_id = st.sec_id  WHERE st.st_status = " + Replace(rsCAgent__MMColParam, "'", "''") + " and s.sec_code = "SE004"  ORDER BY st.st_id DESC"
rsCAgent.CursorType = 0
rsCAgent.CursorLocation = 2
rsCAgent.LockType = 1
rsCAgent.Open()

rsCAgent_numRows = 0
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
    <td> 
      <!-- HEADER STARTS -->
      <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign="bottom"><script>doDate();</script> </td>
          <td width="546" align="right"><img src="images/hdr_logo_01.gif" width="546" height="59"></td>
        </tr>
      </table>
      <!-- HEADER ENDS -->
    </td>
  </tr>
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td>
    <td> 
      <!-- TOPNAV STARTS -->
      <table width="750" height="26"  border="0" cellpadding="0" cellspacing="0" background="images/tnav_bg.gif">
        <tr> 
          <td> <table width="750" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="10"><img src="images/spacer.gif" width="10" height="26"></td>
                <td width="60" align="center"><a href="javascript:void(0);" class="tnav01">Home</a></td>
                <td width="80" align="center"><a href="javascript:void(0);" class="tnav01">About 
                  Us</a> </td>
                <td width="170" align="center"><a href="javascript:void(0);" class="tnav01">Portfolio 
                  of Sold Properties</a> </td>
                <td width="150" align="center"><a href="javascript:void(0);" class="tnav01">Featured 
                  Properties</a> </td>
                <td width="180" align="center"><a href="javascript:void(0);" class="tnav01">Schedule 
                  an Appointment</a> </td>
                <td align="center"><a href="javascript:void(0);" class="tnav01">Search 
                  MLS</a> </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <!-- TOPNAV ENDS -->
    </td>
  </tr>
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td>
    <td><img src="images/tnav_bottom.gif" width="750" height="6"></td>
  </tr>
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td>
    <td valign="top"><table width="720" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="5" valign="top" bgcolor="#FFFFFF"><img src="images/lnav_top_sellers_01.gif" name="lnav_top" width="183" height="21" border="0" usemap="#SELLERS_MAP" id="lnav_top"></td>
          <td valign="top" bgcolor="#FFFFFF"><img src="images/cnt_topShadow.gif" width="566" height="10"></td>
          <td width="1" rowspan="2" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="10"></td>
        </tr>
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="1" valign="top" bgcolor="#E3C66E"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="178" valign="top" bgcolor="#E3C66E"><!-- #BeginLibraryItem "/Library/lnav_buyers_01.lbi" --><table width="178" border="0" cellpadding="0" cellspacing="0" style="height: 100%;"> 
              <tr> 
                <td><img src="images/lnav_ttl_featuredProperties.gif" width="179" height="35"></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="#">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <tr> 
                <td class="lnav01"><a href="javascript:void(0);">Lorem ipsum dolor</a></td> 
              </tr> 
              <!-- LAST LEFTNAV ROW STARTS - DO NOT USE--> 
              <!-- This row sets the bottom border for last link cell above --> 
              <tr> 
                <td class="lnavBorder">&nbsp;</td> 
              </tr> 
              <!-- LAST LEFTNAV ROW END--> 
              <tr> 
                <td><img src="images/lnav_ttl_prepareYourself.gif" width="179" height="35"></td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Getting Ready to Buy</a> </td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Additional Information</a> </td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Industry News</a> </td> 
              </tr> 
              <tr> 
                <td>&nbsp;</td> 
              </tr> 
              <tr style="height: 100%;"> 
                <td><img id="glu" src="images/spacer.gif" width="1" height="1"></td> 
              </tr> 
              <tr> 
                <td><img src="images/lnav_ttl_quickMenu.gif" width="179" height="35"></td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Home</a></td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">About Us</a> </td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Portfolio</a></td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Schedule an Appointment </a></td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Featured Properties </a></td> 
              </tr> 
              <tr> 
                <td class="lnav02"><a href="javascript:void(0);">Search MLS</a> </td> 
              </tr> 
              <tr> 
                <td class="lnav02">&nbsp;</td> 
              </tr> 
            </table><!-- #EndLibraryItem --></td>
          <td width="1" valign="top" bgcolor="#DCB93F"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="1" valign="top" bgcolor="#B06021"><img src="images/spacer.gif" width="1" height="1"></td>
          <td valign="top" bgcolor="#FFFFFF"> 
            <!-- CONTENT AREA STARTS -->
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="images/spacer.gif" width="20" height="68"></td>
                  <td width="400"><img src="images/pg_ttl_PLACEHOLDER.gif" width="400" height="68"> 
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
                  <td colspan="2"> <% If Not rsCAgent.EOF Or Not rsCAgent.BOF Then %>
                    <table width="100%" border="0" cellpadding="2" cellspacing="3">
                      <tr> 
                        <td colspan="2" class="cntHdr01"><%=(rsCAgent.Fields.Item("st_title").Value)%></td>
                      </tr>
                      <tr valign="top"> 
                        <td width="83%"> <p><%=(rsCAgent.Fields.Item("st_desc").Value)%></p>
                          
                        </td>
                        <td width="17%" align="right"><% if (rsCAgent.Fields.Item("st_image").Value) <> "" then%> <img src="imagesDB/<%=(rsCAgent.Fields.Item("st_image").Value)%>"> 
                          <% end if %></td>
                      </tr>
                    </table>
                    <% End If ' end Not rsCAgent.EOF Or NOT rsCAgent.BOF %> </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td>&nbsp;</td>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td>&nbsp;</td>
                  <td colspan="2"> 
                    <!--REPEAT REGION STARTS-->
                    <!--REPEAT REGION ENDS-->
                  </td>
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
                <td width="40%" valign="top" class="nws02"><!-- #BeginLibraryItem "/Library/nws_sample.lbi" --><!--#include file="includes/news.asp" --><!-- #EndLibraryItem --></td>
                <td width="1"><img src="images/nws_divider.gif" width="2" height="100"></td>
                <td width="60%"><!-- #BeginLibraryItem "/Library/tst_sample.lbi" --><!--#include file="includes/testimonial.asp" --><!-- #EndLibraryItem --></td>
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
<map name="BUYERS_MAP" id="BUYERS_MAP">
  <area shape="rect" coords="93,2,181,20" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_buyers_02.gif',0)" onMouseOut="MM_swapImgRestore()">
</map>
<map name="SELLERS_MAP" id="SELLERS_MAP">
  <area shape="rect" coords="2,1,90,19" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_sellers_02.gif',0)" onMouseOut="MM_swapImgRestore()">
</map>
<map name="ABOUT_MAP" id="ABOUT_MAP">
  <area shape="rect" coords="93,2,181,20" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_buyers_02.gif',0)" onMouseOut="MM_swapImgRestore()">
</map>
</body>
</html>
<%
rsCAgent.Close()
Set rsCAgent = Nothing
%>
