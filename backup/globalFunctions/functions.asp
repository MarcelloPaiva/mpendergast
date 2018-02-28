<%
	function aspas(text)
		if text <> "" then aspas = replace(text, "'", "''")
	end function
	
	function getForm(text)
		dim x 
		x = request.Form(text)
		x = trim(x)
		getForm = x
	end function
	
	function getQuery(text)
		dim x 
		x = request.QueryString(text)
		x = trim(x)
		getQuery = x
	end function

	function ValidaNumero(text)
		dim x
		x = trim(text)
		ValidaNumero = false
		if x <> "" and IsNumeric(x) then
			ValidaNumero = true
		end if
	end function

	sub fcErro(text)
		dim x
		x = trim(text)
		x = server.URLEncode(x)
		response.Redirect("erro.asp?erro=" & x)
	end sub

	function getComboStatus(x)
		dim html
		html = html & "<option value='1'"
		if x then html = html & " selected"
		html = html & ">Active</option>"

		html = html & "<option value='0'"
		if not x then html = html & " selected"
		html = html & ">Not Active</option>"

		getComboStatus = html
	end function

	function ValidaEmail(text)
		dim x, strArr, strDot, valida
		x = trim(text)
		strArr = instr(x,"@")
		valida = true
		
		if strArr < 2 then 
			valida = false
		else
			x = mid(x,strArr + 1,len(x))
			strDot = instr(x,".")

			if instr(x,"@") <> 0 then valida = false
			if strDot < 2 then valida = false
			x = mid(x,strDot + 1,len(x))
			if trim(x) = "" then valida = false
		end if
		
		ValidaEmail	= valida
	end function

	function Meses(text)
		dim x, arrMeses
		x = text
		if not ValidaNumero(x) then
			Meses = ""
		else
			x = cint(x)
			if x < 1 or x > 12 then
				Meses = ""
			else
				arrMeses = Array("","January","February","March","April","May","June","July","August","September","Octuber","November","December")
				Meses = arrMeses(x)
			end if
		end if
	end function
	
	function GetComboDay(d)
	dim data, dia, i, msg
		msg = ""
		data = trim(d)

		if data <> "" and isDate(data) then
			dia = day(data)
		else
			dia = 0
		end if
		
		for i = 1 to 31
			if i = dia then
				msg = msg & "<option value='" & i & "' selected>" & right("00" & i,2) & "</option>"
			else
				msg = msg & "<option value='" & i & "'>" & right("00" & i,2) & "</option>"
			end if
		next
		
		GetComboDay = msg
	end function

	function GetComboMonth(d)
	dim data, mes, i, msg
		msg = ""
		data = trim(d)

		if data <> "" and isDate(data) then
			mes = month(data)
		else
			mes = 0
		end if
		
		for i = 1 to 12
			if i = mes then
				msg = msg & "<option value='" & i & "' selected>" & Meses(i) & "</option>"
			else
				msg = msg & "<option value='" & i & "'>" & Meses(i) & "</option>"
			end if
		next
		
		GetComboMonth = msg
	end function

	function getOnOffButton(id, intStatus, section)
		dim return
		
		if intStatus then 
			return = "<img src='images/on.gif' width='21' height='21' style='cursor:hand;' onClick=""document.location='mngStatus.asp?url=" & request.ServerVariables("SCRIPT_NAME") & "?" & request.QueryString() & "&section=" & section & "&status=0&id=" & id & "'"">"
		else
			return = "<img src='images/off.gif' width='21' height='21' style='cursor:hand;' onClick=""document.location='mngStatus.asp?url=" & request.ServerVariables("SCRIPT_NAME") & "?" & request.QueryString() & "&section=" & section & "&status=1&id=" & id & "'"">"
		end if
		
		getOnOffButton = return
	end function
	
	function changeImage(text)
	dim firstletter, return
		if len(text) > 0 then
			firstletter = ucase(mid(text, 1, 1))
			if asc(firstletter) >= 65 and asc(firstletter) <= 91 then 
				return = "<img src='images/letter_" & firstletter & ".gif' align='left'>" & mid(text,2,len(text))
			else
				return = text
			end if
		end if

		changeImage = return
	end function
%>