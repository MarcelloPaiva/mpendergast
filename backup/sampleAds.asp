<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/clsConnection.asp" -->
<%
Dim rsSample__MMColParam
rsSample__MMColParam = "true"
If (Request("MM_EmptyValue") <> "") Then 
  rsSample__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsSample
Dim rsSample_numRows

Set rsSample = Server.CreateObject("ADODB.Recordset")
rsSample.ActiveConnection = MM_clsConnection_STRING
rsSample.Source = "SELECT sa_desc, sa_doc, sa_image, sa_title, sa_url FROM tb_sample_ad WHERE sa_status = " + Replace(rsSample__MMColParam, "'", "''") + " ORDER BY sa_title ASC"
rsSample.CursorType = 0
rsSample.CursorLocation = 2
rsSample.LockType = 1
rsSample.Open()

rsSample_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsSample_numRows = rsSample_numRows + Repeat1__numRows
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
          <td colspan="5" valign="top" bgcolor="#E3C66E"><!--#include file="includes/menuTop_about.asp" --></td>
          <td valign="top" bgcolor="#FFFFFF"><img src="images/cnt_topShadow.gif" width="566" height="10"></td>
          <td width="1" rowspan="2" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="10"></td>
        </tr>
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="1" valign="top" bgcolor="#E3C66E" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="178" valign="top" bgcolor="#E3C66E">
            <!--#include file="includes/menu_about.asp" -->
          </td>
          <td width="1" valign="top" bgcolor="#DCB93F" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td width="1" valign="top" bgcolor="#B06021" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td>
          <td valign="top" bgcolor="#FFFFFF"> 
            <!-- CONTENT AREA STARTS -->
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="images/spacer.gif" width="20" height="68"></td>
                  <td width="400"><img src="images/pg_ttl_sampleAds.gif" width="400" height="68"> 
                    <!--PAGE TITLE -->
                  </td>
                  <td width="125"><img src="images/cnt_topLogo.gif" width="125" height="68"></td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="images/spacer.gif" width="10" height="25"></td>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <% 
While ((Repeat1__numRows <> 0) AND (NOT rsSample.EOF)) 
%>
                <tr bgcolor="#FFFFFF"> 
                  <td>&nbsp;</td>
                  <td colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td class="cntHdr01"><%=(rsSample.Fields.Item("sa_title").Value)%></td>
                        <td width="20">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="cntTxt01"><% if (rsSample.Fields.Item("sa_image").Value) <> "" then %>
                          <img src="imagesDB/<%=(rsSample.Fields.Item("sa_image").Value)%>" hspace="10" vspace="5" align="left">
                        <% end if %> <%=(rsSample.Fields.Item("sa_desc").Value)%></td>
                        <td>&nbsp;</td>
                      </tr>
<% if (rsSample.Fields.Item("sa_url").Value) <> "" then%>
                      <tr> 
                        <td align="right"><a href="<%=(rsSample.Fields.Item("sa_url").Value)%>" target="_blank"><%=(rsSample.Fields.Item("sa_url").Value)%></a></td>
                        <td>&nbsp;</td>
                      </tr>
<% end if %>
<% if (rsSample.Fields.Item("sa_doc").Value) <> "" then %>
                      <tr> 
                        <td align="right">[<a href="imagesDB/<%=(rsSample.Fields.Item("sa_doc").Value)%>" target="_blank">Download Document</a>]</td>
                        <td>&nbsp;</td>
                      </tr>
<% end if %>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsSample.MoveNext()
Wend
%>
                
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
<map name="BUYERS_MAP" id="BUYERS_MAP">
  <area shape="rect" coords="93,2,181,20" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_buyers_02.gif',0)" onMouseOut="MM_swapImgRestore()">
</map>
<map name="SELLERS_MAP" id="SELLERS_MAP">
  <area shape="rect" coords="2,1,90,19" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_sellers_02.gif',0)" onMouseOut="MM_swapImgRestore()">
</map>
</body>
</html>
<%
rsSample.Close()
Set rsSample = Nothing
%>
