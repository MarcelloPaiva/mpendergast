<%
	function menuProperty(id)
		dim html

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyDetails.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0" 
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "detail") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Property Details</a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyPhotos.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "photos") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Photo Gallery </a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyVtour.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "vtour") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Virtual Tour </a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='propertyAddInfo.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "addinfo") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Additional Information </a></td> "
		html = html & "</tr> "

		html = html & "<tr> "
		html = html & "<td class='lnav01sub'><a href='contact.asp?id=" & id & "'><img src='images/lnav_sub_bullet_0"
		if instr(lcase(request.ServerVariables("SCRIPT_NAME")), "contact") <> 0 then 
			html = html & "2"
		else 
			html = html & "1"
		end if	
		html = html & ".gif' width='12' height='18' border='0' align='absmiddle'>Schedule an Appointment </a></td> "
		html = html & "</tr> "

		menuProperty = html
	end function
%>