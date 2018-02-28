<!--#include file="globalFunctions/noCache.asp" -->
<!--#include file="globalFunctions/functions.asp" -->
<!--#include file="globalFunctions/emails.asp" -->
<!--#include file="classes/clsConnection.asp" -->
<!--#include file="classes/clsState.asp" -->
<%
	if request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim obj, pro_name, emailTO, intname, email, phone, address01, address02, stateID, city, zip, desc, erro, msg, sbjct
		'emailTO = "danilo@axeweb.com.br"
		
		pro_name = getForm("pro_name")
		emailTO = getForm("emailTO")
		intName = getForm("name")
		email = getForm("email")
		phone = getForm("phone")
		address01 = getForm("address01")
		address02 = getForm("address02")
		stateID = getForm("state")
		city = getForm("city")
		zip = getForm("zip")
		desc = getForm("desc")
		if pro_name <> "" then sbjct = getForm("pro_name") & " APPOINTMENT REQUEST" 
		if pro_name = "" then 
			sbjct = getForm("assunto")
		end if
		
		if intName = "" then erro = erro & "Fill the name.<br>"
		if not validaEmail(email) then erro = erro & "E-mail is invalid.<br>"
		
		if erro <> "" then
			fcErro(erro)
		end if
		
		set obj = new clsState
		if validanumero(stateID) then obj.setID(stateID)
		obj.fndState
		stateID = obj.getName
		set obj = nothing
		
		msg = "<font face='Tahoma' size='2'>"
		if pro_name <> "" then msg = msg & "Property: " & pro_name & "<br><br>"
		msg = msg & "Name:<b> " & intName & "</b><br>"
		msg = msg & "E-mail: " & email & "<br>"
		msg = msg & "Phone: " & phone & "<br>"
		msg = msg & "Address:<br> " & address01 & "<br>"
		msg = msg & address02 & "<br>"
		msg = msg & "City: " & city & "<br>"
		msg = msg & "State: " & stateID & "<br>"
		msg = msg & "Zip: " & zip & "<br>"
		msg = msg & "Message: " & desc & "<br>"
		msg = msg & "</font>"

		send_mail email , emailTO , sbjct , msg		

	
		response.Redirect("ContactDone.asp")
	end if
	
	response.Redirect("default.asp")	
%>