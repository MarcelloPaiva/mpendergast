<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/clsConnection.asp" -->
<!--#include file="classes/clsState.asp" -->
<%
Dim rsImage__MMColParam
rsImage__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  rsImage__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim rsImage
Dim rsImage_numRows

Set rsImage = Server.CreateObject("ADODB.Recordset")
rsImage.ActiveConnection = MM_clsConnection_STRING
rsImage.Source = "SELECT pro_id, pro_img1, pro_name FROM tb_property WHERE pro_id = " + Replace(rsImage__MMColParam, "'", "''") + ""
rsImage.CursorType = 0
rsImage.CursorLocation = 2
rsImage.LockType = 1
rsImage.Open()

rsImage_numRows = 0
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
    <td>    <!-- TOPNAV STARTS -->    <table width="750" height="26"  border="0" cellpadding="0" cellspacing="0" background="images/tnav_bg.gif">
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
    </table>    <!-- TOPNAV ENDS -->    </td> 
  </tr> 
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td> 
    <td><img src="images/tnav_bottom.gif" width="750" height="6"></td> 
  </tr> 
  <tr> 
    <td width="1"><img src="images/spacer.gif" width="1" height="1"></td> 
    <td valign="top"><table width="720" border="0" cellspacing="0" cellpadding="0"> 
        <tr> 
          <td colspan="5" valign="top" bgcolor="#FFFFFF"><img src="images/lnav_top_about_01.gif" name="lnav_top" width="183" height="21" border="0" usemap="#ABOUT_MAP" id="lnav_top"></td> 
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
                  <td width="400"><img src="images/pg_ttl__scheduleAnAppointment.gif" width="400" height="68">                    <!--PAGE TITLE --> </td> 
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
                        <td> 
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="110" rowspan="2" align="left">
							  <%if not rsImage.eof then%>
							  <img src="imagesDB/<%=(rsImage.Fields.Item("pro_img1").Value)%>" width="110" height="78" hspace="10" border="1">
							  <%end if%></td>
                              <td valign="top"><strong><%if not rsImage.eof then response.write (rsImage.Fields.Item("pro_name").Value)%></strong></td>
                            </tr>
                            <tr>
                              <td valign="top" class="cntTxt01">If you are interested in this or any other property for sale, or if you need additional information, please fill in the contact information below. We would like to help you find the place of your dreams.<br></td>
                            </tr>
                            <tr>
                              <td align="left">&nbsp;</td>
                              <td valign="top" class="cntTxt01">&nbsp;</td>
                            </tr>
                            <tr>
                              <td colspan="2" align="left" class="cntTxt01">You can always reach us at 


 (508) 548-4040 or via fax (508) 540-3721.<br></td>
                            </tr>
                          </table>
<form name="form1" method="post" action="contactAction.asp" onSubmit="return validaForm();">
  <input name="emailTO" type="hidden" value="sales@marthapendergastrealestate.com">
                            <input name="pro_id" type="hidden" id="pro_id" value="<%if not rsImage.eof then response.write (rsImage.Fields.Item("pro_id").Value)%>">
                            <input name="pro_name" type="hidden" id="pro_name" value="<%if not rsImage.eof then response.write(rsImage.Fields.Item("pro_name").Value)%>">
                            <table width="100%" border="0" cellspacing="3" cellpadding="2">
                              <tr> 
                                <td width="18%">Name:*</td>
                                <td width="82%"><input name="name" type="text" class="InputText" id="name" size="40"></td>
                              </tr>
                              <tr> 
                                <td>E-mail:*</td>
                                <td><input name="email" type="text" class="InputText" id="email" size="40"></td>
                              </tr>
                              <tr> 
                                <td>Phone:</td>
                                <td><input name="phone" type="text" class="InputText" id="phone" size="15"></td>
                              </tr>
                              <tr> 
                                <td>Address 01:</td>
                                <td><input name="address01" type="text" class="InputText" id="address01" size="40"></td>
                              </tr>
                              <tr> 
                                <td>Address 02:</td>
                                <td><input name="address02" type="text" class="InputText" id="address02" size="40"></td>
                              </tr>
                              <tr> 
                                <td>State:</td>
                                <td><select name="state" class="InputText" id="state">
                                    <option>Select</option>
<%
	set objState = new clsState
	response.Write objState.getComboState("")
	set objState = nothing
%>                                  </select></td>
                              </tr>
                              <tr> 
                                <td>City:</td>
                                <td><input name="city" type="text" class="InputText" id="city" size="40"></td>
                              </tr>
                              <tr> 
                                <td>ZIP:</td>
                                <td><input name="zip" type="text" class="InputText" id="zip" size="15"></td>
                              </tr>
                              <tr> 
                                <td valign="top">Message:</td>
                                <td><textarea name="desc" cols="50" rows="7" class="InputText" id="desc"></textarea></td>
                              </tr>
                              <tr align="center">
                                <td colspan="2">&nbsp;</td>
                              </tr>
                              <tr align="center"> 
                                <td colspan="2"> <input name="Submit" type="submit" class="Button" value="    Submit    ">                                </td>
                              </tr>
                            </table>
                          </form>
                          <p>&nbsp;</p>
                        </td>
                      <td width="20">&nbsp;</td>
                    </tr>
                  </table></td>
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
<br> 
<br> 
 <map name="BUYERS_MAP" id="BUYERS_MAP"> 
  <area shape="rect" coords="93,2,181,20" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_buyers_02.gif',0)" onMouseOut="MM_swapImgRestore()"> 
 </map>
<map name="SELLERS_MAP" id="SELLERS_MAP"> 
  <area shape="rect" coords="2,1,90,19" href="#" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_sellers_02.gif',0)" onMouseOut="MM_swapImgRestore()"> 
</map> 
<map name="ABOUT_MAP" id="ABOUT_MAP"> 
  <area shape="rect" coords="93,2,181,20" href="genSections.asp?sec=2" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_about_03.gif',0)" onMouseOut="MM_swapImgRestore()"> 
  <area shape="rect" coords="1,2,91,20" href="genSections.asp?sec=1" onMouseOver="MM_swapImage('lnav_top','','images/lnav_top_about_02.gif',0)" onMouseOut="MM_swapImgRestore()">
</map>  

</body>
</html>
<%
rsImage.Close()
Set rsImage = Nothing
%>
