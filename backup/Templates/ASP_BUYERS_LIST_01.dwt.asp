<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/clsConnection.asp" -->
<%
Dim rsProperties__pStat
rsProperties__pStat = "true"
If (Request("MM_EmptyValue") <> "") Then 
  rsProperties__pStat = Request("MM_EmptyValue")
End If
%>
<%
Dim rsProperties
Dim rsProperties_numRows

Set rsProperties = Server.CreateObject("ADODB.Recordset")
rsProperties.ActiveConnection = MM_clsConnection_STRING
rsProperties.Source = "SELECT *  FROM tb_property  WHERE pro_status = " + Replace(rsProperties__pStat, "'", "''") + " AND pro_sellDate is null  ORDER BY pro_name ASC"
rsProperties.CursorType = 0
rsProperties.CursorLocation = 2
rsProperties.LockType = 1
rsProperties.Open()

rsProperties_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<!-- TemplateBeginEditable name="doctitle" -->
<title>Martha Pendergast Real Estate -</title>
<!-- TemplateEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2" src="../js/date.js" type="text/javascript"></script>
<script language="JavaScript1.2" src="../js/global.js" type="text/javascript"></script>
<link href="../css/style_01.css" rel="stylesheet" type="text/css">
<!-- TemplateBeginEditable name="head" --><!-- TemplateEndEditable -->
</head>
<body onLoad="setHeight();MM_preloadImages('../images/ftr_terms_01.gif')"> 
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0"> 
  <tr> 
    <td width="1"><img src="../images/spacer.gif" width="1" height="1"></td> 
    <td> <!-- HEADER STARTS --> 
      <table width="750" border="0" cellspacing="0" cellpadding="0"> 
        <tr> 
          <td valign="bottom"><script>doDate();</script> </td> 
          <td width="546" align="right"><img src="../images/hdr_logo_01.gif" width="546" height="59"></td> 
        </tr> 
      </table> 
      <!-- HEADER ENDS --> </td> 
  </tr> 
  <tr> 
    <td width="1"><img src="../images/spacer.gif" width="1" height="1"></td> 
    <td><!-- TemplateBeginEditable name="GLOBAL NAVIGATION" -->
    <!-- TOPNAV STARTS -->
    <table width="750" height="26"  border="0" cellpadding="0" cellspacing="0" background="../images/tnav_bg.gif">
      <tr>
        <td><!-- #BeginLibraryItem "/Library/global_nav_01.lbi" --><table width="750" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="10"><img src="../images/spacer.gif" width="10" height="26"></td>
              <td width="60" align="center"><a href="../default.asp" class="tnav01">Home</a></td>
              <td width="80" align="center"><a href="../genSections.asp?sec=5" class="tnav01">About Us</a> </td>
              <td width="170" align="center"><a href="../portfolio.asp" class="tnav01">Portfolio of Sold Properties</a> </td>
              <td width="150" align="center"><a href="../propertyList.asp" class="tnav01">Featured Properties</a> </td>
              <td width="180" align="center"><a href="../contact2.asp" class="tnav01">Schedule an Appointment</a> </td>
              <td align="center"><a href="../searchMLS.asp" class="tnav01">Search MLS</a> </td>
            </tr>
        </table><!-- #EndLibraryItem --></td>
      </tr>
    </table>
    <!-- TOPNAV ENDS -->
    <!-- TemplateEndEditable --></td> 
  </tr> 
  <tr> 
    <td width="1"><img src="../images/spacer.gif" width="1" height="1"></td> 
    <td><img src="../images/tnav_bottom.gif" width="750" height="6"></td> 
  </tr> 
  <tr> 
    <td width="1"><img src="../images/spacer.gif" width="1" height="1"></td> 
    <td valign="top"><table width="720" border="0" cellspacing="0" cellpadding="0"> 
        <tr> 
          <td colspan="5" valign="top" bgcolor="#FFFFFF"><img src="../images/lnav_top_buyers_01.gif" name="lnav_top" width="183" height="21" border="0" usemap="#Map" id="lnav_top"></td> 
          <td valign="top" bgcolor="#FFFFFF"><img src="../images/cnt_topShadow.gif" width="566" height="10"></td> 
          <td width="1" rowspan="2" valign="top" bgcolor="#4C0000"><img src="../images/spacer.gif" width="1" height="10"></td> 
        </tr> 
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000"><img src="../images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#E3C66E"><img src="../images/spacer.gif" width="1" height="1"></td> 
          <td width="178" valign="top" bgcolor="#E3C66E"><!-- TemplateBeginEditable name="LEFT NAVIGATION" --><!-- TemplateEndEditable --></td> 
          <td width="1" valign="top" bgcolor="#DCB93F"><img src="../images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#B06021"><img src="../images/spacer.gif" width="1" height="1"></td> 
          <td valign="top" bgcolor="#FFFFFF"> <!-- CONTENT AREA STARTS --> 
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="../images/spacer.gif" width="20" height="68"></td> 
                  <td width="400"><!-- TemplateBeginEditable name="PAGE TITLE" --><img src="../images/pg_ttl_PLACEHOLDER.gif" width="400" height="68"><!-- TemplateEndEditable -->                    <!--PAGE TITLE --> </td> 
                  <td width="125"><img src="../images/cnt_topLogo.gif" width="125" height="68"></td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="../images/spacer.gif" width="10" height="25"></td> 
                  <td colspan="2">&nbsp;</td> 
                </tr> 
                <tr bgcolor="#FFFFFF">
                  <td>&nbsp;</td>
                  <td colspan="2"><!-- TemplateBeginEditable name="CONTENT AREA" -->
                  <!--REPEAT REGION STARTS-->
                  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td colspan="4" class="cntHdr01"><a href="javascript:void(0);">Property Title</a></td>
                    </tr>
                    <tr>
                      <td colspan="4" class="cntHdr01"><img src="../images/spacer.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td width="160" valign="top"><a href="javascript:void(0);"><img src="../images/prp_pht_PLACEHOLDER.gif" alt="Click to view property details." width="160" height="160" border="0"></a><img src="../images/prp_pht_viewDetails.gif" width="160" height="13"></td>
                      <td width="10" valign="top"><img src="../images/spacer.gif" width="10" height="10"></td>
                      <td valign="top" class="cntTxt01">Lorem ipsum dolor sit amet, consectetaur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum Et harumd und lookum like Greek to me, dereud facilis ( <a href="javascript:void(0);">500 characters place holder</a> )</td>
                      <td width="10" class="cntTxt01"><img src="../images/spacer.gif" width="20" height="10"></td>
                    </tr>
                    <tr>
                      <td valign="top"><img src="../images/spacer.gif" width="10" height="20"></td>
                      <td valign="top">&nbsp;</td>
                      <td valign="top">&nbsp;</td>
                      <td valign="top">&nbsp;</td>
                    </tr>
                  </table>
                  <!--REPEAT REGION ENDS-->
                  <!-- TemplateEndEditable --></td>
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
          <td width="1" valign="top" bgcolor="#4C0000"><img src="../images/spacer.gif" width="1" height="1"></td> 
          <td valign="top" bgcolor="#FFFFFF" class="nws01"><table width="100%"  border="0" cellspacing="0" cellpadding="0"> 
              <tr> 
                <td height="25"><img src="../images/nws_ttl_newsflash.gif" width="178" height="25"></td> 
                <td width="1"><img src="../images/spacer.gif" width="1" height="1"></td> 
                <td><a href="javascript:;" onMouseOver="MM_swapImage('nws_ttl_testimonials','','../images/nws_ttl_testimonials_02.gif',0)" onMouseOut="MM_swapImgRestore()"><img src="../images/nws_ttl_testimonials.gif" name="nws_ttl_testimonials" width="208" height="25" border="0" id="nws_ttl_testimonials"></a></td> 
              </tr> 
              <tr> 
                <td width="40%" valign="top" class="nws02"><!-- TemplateBeginEditable name="NEWSFLASH" --><!-- #BeginLibraryItem "/Library/nws_sample.lbi" -->
                <script language="JavaScript" src="../js/newsflash.js"></script>
                <!-- #EndLibraryItem --><!-- TemplateEndEditable --></td> 
                <td width="1"><img src="../images/nws_divider.gif" width="2" height="100"></td> 
                <td width="60%"><!-- TemplateBeginEditable name="TESTIMONIALS" --><!-- #BeginLibraryItem "/Library/tst_sample.lbi" --><!--#include file="../Library/includes/testimonial.asp" --><!-- #EndLibraryItem --><!-- TemplateEndEditable --></td> 
              </tr> 
              <tr> 
                <td>&nbsp;</td> 
                <td><img src="../images/spacer.gif" width="1" height="10"></td> 
                <td>&nbsp;</td> 
              </tr> 
            </table></td> 
          <td width="1" valign="top" bgcolor="#4C0000"><img src="../images/spacer.gif" width="1" height="1"></td> 
        </tr> 
        <tr> 
          <td valign="top" bgcolor="#4C0000"><img src="../images/spacer.gif" width="1" height="15"></td> 
          <td valign="top" bgcolor="#FFFFFF" class="ftr01"><!-- TemplateBeginEditable name="FOOTER" --><img src="../images/ftr_legal_01.gif" name="ftr_legal" width="720" height="13" border="0" usemap="#ftr_legal" id="ftr_legal"><!-- TemplateEndEditable --></td> 
          <td valign="top" bgcolor="#4C0000"><img src="../images/spacer.gif" width="1" height="1"></td> 
        </tr> 
      </table></td> 
  </tr> 
</table> 
<map name="Map"> 
  <area shape="rect" coords="93,2,181,20" href="#" onMouseOver="MM_swapImage('lnav_top','','../images/lnav_top_buyers_02.gif',0)" onMouseOut="MM_swapImgRestore()"> 
</map> 
 
<br> 
<br> 
</body>
</html>
<%
rsProperties.Close()
Set rsProperties = Nothing
%>
