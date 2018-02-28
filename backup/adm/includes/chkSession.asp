<%
	if session("USR_" & session.SessionID) then
		dim iContUser, arrUserAccess, chkUserStatus
		chkUserStatus = false
		arrUserAccess = session("usr_arrAccess")
		
		if isArray(arrUserAccess) then
			for iContUser = 0 to ubound(arrUserAccess,2)
				if arrUserAccess(0,iContUser) = ("AC" & right("000" & section_code ,3)) then
					chkUserStatus = true
				end if
			next
		end if
				
	if not chkUserStatus then fcErro("Permission Denied. <br>You don't have access in this section.")
	else
		response.Redirect("default.asp")
	end if
%>