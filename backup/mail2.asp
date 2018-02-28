<%@LANGUAGE="VBSCRIPT"%>
<%
Dim MailerPath
Dim FromEmail
Dim ToEmail
Dim FromName
Dim Subject
Dim fieldName
Dim fieldValue
Dim Message
Dim mailRedirect
Dim mailComp
MailerPath = "mail.marthapendergastrealestate.com"
FromEmail = "muqueca@hotmail.com"
ToEmail="paiva@comcast.net"
FromName="my Name"
Subject="my subject"
For i = 1 to Request.Form.Count
		fieldName = Request.Form.Key(i)
		fieldValue = Request.Form.Item(i)
		Message = Message & fieldName & ": " & fieldValue & VbCrLf
	Next
mailRedirect="ContactDone.asp"
mailComp=1
If (cStr(Request("Submit")) <> "") Then
Select case mailComp
case 1
	call CDONTS_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
case 2
	call ASPMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
case 3 
	call ASPQMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
case 4
	call JMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
case 5 
	call ASPEMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
case 6 
	call SASmtpMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
case else
	Response.Write("We are sorry, the system has encountered an error")
End Select
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/ASP_BUYERS_LIST_01.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Martha Pendergast Real Estate -</title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2" src="js/date.js" type="text/javascript"></script>
<script language="JavaScript1.2" src="js/global.js" type="text/javascript"></script>
<link href="css/style_01.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
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
    <td><!-- InstanceBeginEditable name="GLOBAL NAVIGATION" -->
    <!-- TOPNAV STARTS -->
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
    <!-- TOPNAV ENDS -->
    <!-- InstanceEndEditable --></td> 
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
          <td width="178" valign="top" bgcolor="#E3C66E"><!-- InstanceBeginEditable name="LEFT NAVIGATION" --><!-- InstanceEndEditable --></td> 
          <td width="1" valign="top" bgcolor="#DCB93F"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td width="1" valign="top" bgcolor="#B06021"><img src="images/spacer.gif" width="1" height="1"></td> 
          <td valign="top" bgcolor="#FFFFFF"> <!-- CONTENT AREA STARTS --> 
            <div id="content"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
                <tr bgcolor="#FFFFFF"> 
                  <td width="20"><img src="images/spacer.gif" width="20" height="68"></td> 
                  <td width="400"><!-- InstanceBeginEditable name="PAGE TITLE" --><img src="images/pg_ttl_PLACEHOLDER.gif" width="400" height="68"><!-- InstanceEndEditable -->                    <!--PAGE TITLE --> </td> 
                  <td width="125"><img src="images/cnt_topLogo.gif" width="125" height="68"></td> 
                </tr> 
                <tr bgcolor="#FFFFFF"> 
                  <td><img src="images/spacer.gif" width="10" height="25"></td> 
                  <td colspan="2">&nbsp;</td> 
                </tr> 
                <tr bgcolor="#FFFFFF">
                  <td>&nbsp;</td>
                  <td colspan="2"><!-- InstanceBeginEditable name="CONTENT AREA" -->
                  <!--REPEAT REGION STARTS-->
                  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td colspan="4" class="cntHdr01"><a href="javascript:void(0);">Property Title</a></td>
                    </tr>
                    <tr>
                      <td colspan="4" class="cntHdr01"><img src="images/spacer.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td width="160" valign="top"><a href="javascript:void(0);"><img src="images/prp_pht_PLACEHOLDER.gif" alt="Click to view property details." width="160" height="160" border="0"></a><img src="images/prp_pht_viewDetails.gif" width="160" height="13"></td>
                      <td width="10" valign="top"><img src="images/spacer.gif" width="10" height="10"></td>
                      <td valign="top" class="cntTxt01"><p>Lorem ipsum dolor sit amet, consectetaur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum Et harumd und lookum like Greek to me, dereud facilis ( <a href="javascript:void(0);">500 characters place holder</a> )</p>
                        <form name="form1" method="post" action="">
                          <input type="submit" name="Submit" value="Submit">
                        </form>                        <p>&nbsp;</p></td>
                      <td width="10" class="cntTxt01"><img src="images/spacer.gif" width="20" height="10"></td>
                    </tr>
                    <tr>
                      <td valign="top"><img src="images/spacer.gif" width="10" height="20"></td>
                      <td valign="top">&nbsp;</td>
                      <td valign="top">&nbsp;</td>
                      <td valign="top">&nbsp;</td>
                    </tr>
                  </table>
                  <!--REPEAT REGION ENDS-->
                  <!-- InstanceEndEditable --></td>
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
                <td><a href="javascript:;" onMouseOver="MM_swapImage('nws_ttl_testimonials','','images/nws_ttl_testimonials_02.gif',0)" onMouseOut="MM_swapImgRestore()"><img src="images/nws_ttl_testimonials.gif" name="nws_ttl_testimonials" width="208" height="25" border="0" id="nws_ttl_testimonials"></a></td> 
              </tr> 
              <tr> 
                <td width="40%" valign="top" class="nws02"><!-- InstanceBeginEditable name="NEWSFLASH" --><!-- InstanceEndEditable --></td> 
                <td width="1"><img src="images/nws_divider.gif" width="2" height="100"></td> 
                <td width="60%"><!-- InstanceBeginEditable name="TESTIMONIALS" --><!-- #BeginLibraryItem "/Library/tst_sample.lbi" --><!-- InstanceEndEditable --></td> 
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
          <td valign="top" bgcolor="#FFFFFF" class="ftr01"><!-- InstanceBeginEditable name="FOOTER" --><img src="images/ftr_legal_01.gif" name="ftr_legal" width="720" height="13" border="0" usemap="#ftr_legal" id="ftr_legal"><!-- InstanceEndEditable --></td> 
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
<!-- InstanceEnd --></html>
<%
'======================================
' Sends an email with CDONTS 
function CDONTS_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
	Dim Mailer
	Set Mailer = Server.CreateObject("CDO.message") 
	Mailer.To = ToEmail
	Mailer.From = FromEmail
	Mailer.Subject = Subject
	Mailer.TextBody = Message
	Mailer.MailFormat = 1
	Mailer.BodyFormat = 1
	Mailer.Send
	Set Mailer = Nothing
	CDONTS_Mailer = true
	Response.Redirect(mailRedirect)
end function

'======================================
' Sends an email with ASPMail 
function ASPMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
	Dim Mailer
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.RemoteHost = MailerPath
	Mailer.ContentType = "text/plain"
	Mailer.FromName = FromName
	Mailer.FromAddress = FromEmail
	Mailer.AddRecipient ToName, ToEmail
	Mailer.Subject = Subject
	Mailer.BodyText = Message
	Mailer.SendMail
	Set Mailer = Nothing
	ASPMail_Mailer = true
	Response.Redirect(mailRedirect)
end function

'======================================
' Sends an email with ASPQMail 
function ASPQMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.RemoteHost = MailerPath
	Mailer.ContentType = "test/plain"
	Mailer.FromName = FromName
	Mailer.FromAddress = FromEmail
	Mailer.AddRecipient ToName, ToEmail
	Mailer.Subject = Subject
	Mailer.BodyText = Message
	Mailer.QMessage = true
	Mailer.SendMail
	Set Mailer = Nothing
	ASPQMail_Mailer = true
	Response.Redirect(mailRedirect)
end function

'======================================
' Sends an email with SASmtpMail 
function SASmtpMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail")
	Mailer.RemoteHost = MailerPath
	Mailer.contenttype = "text/plain"
	Mailer.AddRecipient ToName, ToEmail
	Mailer.FromName = FromName
	Mailer.FromAddress = FromEmail
	Mailer.Subject = Subject
	Mailer.BodyText = Message
	Mailer.SendMail
	Set Mailer = Nothing
	SASmtpMail_Mailer = true
	Response.Redirect(mailRedirect)
end function

'======================================
' Sends an email with JMail 
function JMail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("JMail.SMTPMail") 
	Mailer.ServerAddress = MailerPath & ":"
	Mailer.contenttype = "text/plain"
	Mailer.AddRecipient ToEmail
	Mailer.Sender = FromEmail
	Mailer.Subject = Subject
	Mailer.TextBody = Message
	Mailer.Execute
	Set Mailer = Nothing
	JMail_Mailer = true
	Response.Redirect(mailRedirect)
end function

'======================================
' Sends an email with ASPEmail 
function ASPEmail_Mailer(Message, FromEmail, ToEmail, FromName, ToName, Subject, MailerPath,mailRedirect)
	on error resume next
	Dim Mailer
	Set Mailer = Server.CreateObject("Persits.MailSender")  
	Mailer.Host = MailerPath
	Mailer.From = FromEmail
	Mailer.FromName = FromName
	Mailer.AddAddress ToName, ToEmail
	Mailer.Subject = Subject
	Mailer.TextBody = Message
	Mailer.Send
	Set Mailer = Nothing
	ASPEmail_Mailer = true
	Response.Redirect(mailRedirect)
end function
%>
