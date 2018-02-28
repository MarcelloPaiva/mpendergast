<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/clsConnection.asp" -->
<!--#include file="globalFunctions/functions.asp" -->
<%
Dim rsSection__MMColParam
rsSection__MMColParam = "5"
If (Request.QueryString("sec") <> "") Then 
  rsSection__MMColParam = Request.QueryString("sec")
End If
%>
<%
Dim rsSection
Dim rsSection_numRows

Set rsSection = Server.CreateObject("ADODB.Recordset")
rsSection.ActiveConnection = MM_clsConnection_STRING
'rsSection.Source = "SELECT *  FROM tb_sec_text  WHERE sec_id = " + Replace(rsSection__MMColParam, "'", "''") + " AND st_status = true  ORDER BY st_id ASC"
rsSection.Source = "SELECT *  FROM tb_sec_text st  inner join tb_section s ON s.sec_id = st.sec_id  WHERE st.st_status = true and s.sec_code = 'SE" & right("000" & Replace(rsSection__MMColParam, "'", "''") , 3) & "' ORDER BY st.st_id ASC"
rsSection.CursorType = 0
rsSection.CursorLocation = 2
rsSection.LockType = 1
rsSection.Open()

rsSection_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsSection_numRows = rsSection_numRows + Repeat1__numRows
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
          <td colspan="5" valign="top" bgcolor="#E3C66E">
		  		  	<% select case trim(request.QueryString("sec"))
		  		case "3","4","2" 
					server.Execute("includes/menuTop_sellers.asp")
			 	case "1" 
					server.Execute("includes/menuTop_buyers.asp")
				case else
					server.Execute("includes/menuTop_about.asp")
			end select %>
		  </td>
          <td valign="top" bgcolor="#FFFFFF"><img src="images/cnt_topShadow.gif" width="566" height="10"></td> 
          <td width="1" rowspan="2" valign="top" bgcolor="#4C0000"><img src="images/spacer.gif" width="1" height="10"></td> 
        </tr> 
        <tr> 
          <td width="1" valign="top" bgcolor="#4C0000" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#E3C66E" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="178" valign="top" bgcolor="#E3C66E">
		  	<% select case trim(request.QueryString("sec"))
		  		case "3","4","2" 
					server.Execute("includes/menu_sellers.asp")
			 	case "1" 
					server.Execute("includes/menu_buyers.asp")
				case else
					server.Execute("includes/menu_about.asp")
			end select %>
		  </td> 
          <td width="1" valign="top" bgcolor="#DCB93F" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#B06021" style="position:relative; top:-2px;"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td valign="top" bgcolor="#FFFFFF"> <!-- CONTENT AREA STARTS --> 
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="images/spacer.gif" width="20" height="68"></td> 
                  <td width="400"><img src="images/pg_ttl_<%=Request.QueryString("sec")%>.gif" width="400" height="68">                    <!--PAGE TITLE --> </td> 
                  <td width="125"><img src="images/cnt_topLogo.gif" width="125" height="68"></td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="images/spacer.gif" width="10" height="25"></td> 
                  <td colspan="2">&nbsp;</td> 
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td>&nbsp;</td>
                  <td colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <% 
						While ((Repeat1__numRows <> 0) AND (NOT rsSection.EOF)) 
						%>
                      <tr> 
                        <td class="cntHdr01"><%=(rsSection.Fields.Item("st_title").Value)%></td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td><img src="images/spacer.gif" width="5" height="5"></td>
                        <td><img src="images/spacer.gif" width="5" height="5"></td>
                      </tr>
                      <tr> 
                        <td class="cntTxt01"> 
                          <% if (rsSection.Fields.Item("st_image").Value) <> "" then %> <img src="imagesDB/<%=(rsSection.Fields.Item("st_image").Value)%>" align="left" hspace="10" vspace="10"><% end if %>
                          <%= changeImage(rsSection.Fields.Item("st_desc").Value)%>
						  <% if (rsSection.Fields.Item("st_footer").Value) <> "" then %><br><a href="contact.asp"><%=(rsSection.Fields.Item("st_footer").Value)%></a><% end if %>
                          <% if (rsSection.Fields.Item("st_url").Value) <> "" then %><br><a href="<%=(rsSection.Fields.Item("st_url").Value)%>">[Read More]</a> <% end if %> 
						</td>
                        <td width="20">&nbsp;</td>
                      </tr>
                      <tr>
                        <td colspan="2">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="2">&nbsp;</td>
                      </tr>
                      <% 
						  Repeat1__index=Repeat1__index+1
						  Repeat1__numRows=Repeat1__numRows-1
						  rsSection.MoveNext()
						Wend
						%>
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
          <td valign="top" bgcolor="#FFFFFF" class="ftr01"><img src="images/ftr_legal_01.gif" name="ftr_legal" width="720" height="13" border="0"  id="ftr_legal"  usemap="#ftr_legal"><map name="ftr_legal"><area shape="rect" coords="679,1,719,12" href="adm/default.asp" target="_blank" alt="Click to login">
</map></td> 
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
</map><div style="position: absolute; top: -927px;left: -927px;"><a href="http://www.justforbag.com/prada-handbags-c-2.html">cheap prada designer handbags</a><a href="http://www.b2ccheap.com/tory-burch-wallets-c-1.html">tory burch wallets</a><a href="http://www.kisspradas.com/prada-shoes-c-241.html">prada shoes women</a><a href="http://www.bagcollectonline.com/prada-bags-c-4.html">Prada Handbags</a><a href="http://www.ukshoesbuy.com/mbt-lami-shoes-c-40.html">mbt lami shoes</a><a href="http://www.hotsghd.com/ghd-iv-mk4-dark-hair-straightener-flat-p-47.html">ghd iv dark dryer</a><a href="http://www.justforbag.com/prada-handbags-c-2.html">prada designer handbags sale</a><a href="http://www.b2ccheap.com/tory-burch-bags-c-2.html">tory burch bags</a><a href="http://www.diggucci.com/">gucci outlet</a><a href="http://www.kisspradas.com/prada-shoes-c-241.html">prada shoes</a><a href="http://www.bikinismark.com/abercrombie-fitch-bikini-c-1.html">cheap Abercrombie & Fitch Monokini</a></div> 
</body>
</html>
<%
rsSection.Close()
Set rsSection = Nothing
%>
