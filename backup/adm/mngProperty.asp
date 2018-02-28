<!--#include file="../globalFunctions/noCache.asp" -->
<%
	dim section_code
	section_code = 15
%>
<!--#include file="includes/chkSession.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim objForm, return
		
		return = ""
		
		set objForm = new clsForms
		return = objForm.frmProperty
		set objForm = nothing
		
		if return <> "" then
			fcErro(return)
		end if
		
		response.Redirect("listProperty.asp")
	end if
	
	dim obj, objP, id, cit_id, cli_id, clientName, pt_id, intName, address, intNumber, desc, price, txt1, txt2, img1, img2, sellDate, intStatus, vtourUrl, vtourDesc
	id = getQuery("id")
	cli_id = getQuery("cli_id")
	intStatus = true
	
	if getQuery("action") = "delete" and validanumero(id) then
		set obj = new clsProperty
		obj.setID(id)
		obj.delProperty
		set obj = nothing
		
		response.Redirect("listProperty.asp")
	end if
	
	if not validaNumero(id) and not ValidaNUmero(cli_id) then response.Redirect("mngClient.asp")
	
	if validaNumero(id) then
		set objP = new clsProperty
		objp.setID(id)
		objP.fndProperty
		cit_id = objP.getCityID()
		pt_id = objP.getPropertyTypeID()
		cli_id = objP.getClientID()
		intName = objP.getname()
		address = objP.getAddress()
		intNumber = objP.getNumber()
		desc = objP.getDesc()
		price = objP.getPrice()
		txt1 = objP.getTxt1()
		txt2 = objP.getTxt2()
		img1 = objP.getImg1()
		img2 = objP.getImg2()
		selldate = objP.getSellDate()
		intStatus = objP.getStatus()
		vtourUrl = objP.getVtourURL()
		vtourDesc = objP.getVtourDesc()
		set objP = nothing
	end if
	
	if validaNumero(cli_id) then
		set obj = new clsClient
		obj.setID(cli_id)
		obj.fndClient
		clientName = obj.getName
		set obj = nothing
	end if
%>
<!--#include file="includes/head.asp" -->
<!--#include file="includes/top.asp" -->
<!--#include file="includes/main.asp" -->
      <form action="<%= request.ServerVariables("SCRIPT_NAME") %>" method="post" enctype="multipart/form-data" name="form1">
        <input name="id" type="hidden" id="id" value="<%= id %>">
		<input name="cli_id" type="hidden" value="<%= cli_id %>">
        <table width="500" border="0" align="center" cellpadding="2" cellspacing="3">
          <tr> 
            <td colspan="2" class="admHdr01">PROPERTY</td>
          </tr>
          <tr>
            <td colspan="2" class="InputText"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="3">
              <tr class="txt2">
                <td class="photo01">Owner (s) </td>
                <td align="center" class="photo01">Features</td>
                <td align="center" class="photo01">Photo Galley </td>
              </tr>
              <tr class="txt2">
                <td width="363"><span class="photo01"><%= clientName %></span> </td>
                <td width="45" align="center"><a href="mngPropertyFeature.asp?pro_id=<%= id %>"><img src="images/features.gif" alt="FEATURES" border="0"></a></td>
                <td width="56" align="center"><a href="listPropertyImage.asp?pro_id=<%= id %>"><img src="images/images.gif" alt="PHOTOS" border="0"></a></td>
              </tr>
            </table></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td class="cntHdr02">&nbsp;</td>
          </tr>
          <tr> 
            <td width="97">Name:</td>
            <td width="386"><input name="name" type="text" id="name" size="40" maxlength="100" value="<%= intName %>" style="font-weight:bold "></td>
          </tr>
          <tr> 
            <td>Type:</td>
            <td><select name="pt_id" id="pt_id">
                <option value="">Select</option>
                <%
	set obj = new clsPropertyType
	response.Write obj.getComboPropertyType(pt_id)
	set obj = nothing
%>
              </select> </td>
          </tr>
          <tr> 
            <td>City</td>
            <td><select name="cit_id" id="cit_id">
                <option value="">Select</option>
                <%
	set obj= new clsCity
	response.Write obj.getComboCity(cit_id)
	set obj = nothing
%>
              </select></td>
          </tr>
          <tr> 
            <td>Address:</td>
            <td><input name="address" type="text" id="address" size="40" maxlength="100" value="<%= address %>"></td>
          </tr>
          <tr> 
            <td>Price:</td>
            <td>$
              <input name="price" type="text" id="price" size="15" maxlength="15" value="<%= FormatNumber( price , -1, -2, -2, -2) %>">
              <span class="photo01">(e.g. 1,000.00)</span></td>
          </tr>
          <tr> 
            <td valign="top">Description:</td>
            <td><textarea name="desc" cols="50" rows="10" id="desc"><%= desc %></textarea></td>
          </tr>
          <tr> 
            <td valign="top">Photo 1:<br>
              <span class="photo01">(200 x 190 pixels)</span></td>
            <td valign="middle"><span class="photo01">Displays on Property List and Property Detail pages.</span> <br>
              <input name="file1" type="file" id="file1"> <% if img1 <> "" then %>
                <a href="view_image.asp?id=<%= id %>&section=property&imgN=1&image=<%= img1 %>&url=<%= request.ServerVariables("SCRIPT_NAME") & "?" & request.QueryString %>"><img src="../imagesDB/<%= img1 %>" width="42" height="25" border="0" align="absmiddle" alt="<%= img1 %>"></a>
            <% end if %></td></tr>
          <tr> 
            <td valign="top">Text 1:</td>
            <td><textarea name="txt1" cols="50" rows="4" id="txt1"><%= txt1 %></textarea></td>
          </tr>
          <tr> 
            <td valign="top">Photo 2:<br>
              <span class="photo01">(200 x 190 pixels)</span></td>
            <td><span class="photo01">Displays on  Property Detail page only.</span><br>
              <input name="file2" type="file" id="file2"> <% if img2 <> "" then %>
              <a href="view_image.asp?id=<%= id %>&section=property&imgN=1&image=<%= img2 %>&url=<%= request.ServerVariables("SCRIPT_NAME") & "?" & request.QueryString %>"><img src="../imagesDB/<%= img2 %>" width="42" height="25" border="0" align="absmiddle" alt="<%= img2 %>"></a>
              <% end if %> </td>
          </tr>
          <tr> 
            <td valign="top">Text 2:</td>
            <td><textarea name="txt2" cols="50" rows="4" id="txt2"><%= txt2 %></textarea></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>Status:</td>
            <td><select name="number" id="number">
                <option value="0"<% if intNumber = 0 then %> selected<% end if %>>For Sale</option>
                <option value="1"<% if intNumber = 1 then %> selected<% end if %>>Under Agreement</option>
                <option value="2"<% if intNumber = 2 then %> selected<% end if %>>Sold</option>
              </select>
                <select name="status" id="status">
                  <%= getComboStatus(intStatus) %>
              </select> 
                <span class="cntTxt01" style="font-size: 9px">(&quot;Sold&quot; will send to Portfolio)</span> </td>
          </tr>
          <tr> 
            <td>Sell Date:</td>
            <td><select name="sMonth" id="sMonth">
                <option value="">Month</option>
                <%= getComboMonth(sellDate) %> </select> <select name="sDay" id="sDay">
                <option value="">Day</option>
                <%= getComboDay(sellDate) %> </select> <input name="sYear" type="text" id="sYear" size="4" maxlength="5" <% if isDate(sellDate) then %>value="<%= year(selldate) %>"<% end if %>> 
                <span class="cntTxt01" style="font-size: 9px">(Will send property to Portfolio)</span> </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>Virtual Tour URL:</td>
            <td><input name="vtourUrl" type="text" id="vtourUrl" size="40" maxlength="200" value="<%= vtourUrl %>"> 
              <span class="cntTxt01" style="font-size: 9px">External Link</span> </td>
          </tr>
          <tr> 
            <td valign="top">Virtual Tour Description:</td>
            <td><textarea name="vtourDesc" cols="50" rows="5" id="vtourDesc"><%= vtourDesc %></textarea></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="right"><a href="listProperty.asp"><img src="../img/bt_cancel.gif" width="63" height="25" border="0"></a> 
              <input name="imageField" type="image" src="../img/bt_salvar.gif" width="63" height="25" border="0"></td>
          </tr>
        </table>
      </form> 
<!--#include file="includes/main_end.asp" -->
<!--#include file="includes/bottom.asp" -->
