<%
	class clsUser
		private conn, rs, intID, intName, intLogin, intPassword, intStatus
		public arrAccess
		
		private sub class_initialize
			intID = -1
			set rs = server.CreateObject("adodb.recordset")
			set conn = new clsConnection
		end sub
		
		private sub class_terminate
			if isObject(rs) then set rs = nothing
			if isObject(conn) then set conn = nothing
		end sub
		
		public sub setID(x)
			intID = x
		end sub
		
		public sub setName(x)
			intName = x
		end sub

		public sub setLogin(x)
			intLogin = x
		end sub
		
		public sub setPassword(x)
			dim crypt
			set crypt = new clsCryptPass
			intPassword = crypt.crypt(x)
			set crypt = nothing
		end sub
		
		public sub setStatus(x)
			intStatus = x
		end sub

		public function getID()
			getID = intID
		end function
		
		public function getName()
			getName = intName
		end function

		public function getLogin()
			getLogin = intLogin
		end function
		
		private function getPassword()
			getPassword = intPassword
		end function
		
		public function getStatus()
			getStatus = intStatus
		end function
		
		public function MngUser()
			dim sql
			
			sql = "select * from tb_user where usr_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("usr_name") = getName
			rs("usr_login") = getLogin
			rs("usr_password") = getPassword
			rs("usr_status") = getStatus
			rs.update
			rs.close
		end function
		
		public function chkUser(login, pass)
			dim sql, c
			
			set c = new clsCryptPass
			sql = "select usr_id, usr_name from tb_user where usr_status = true and usr_login = '" & aspas(login) & "' and usr_password = '" & c.crypt(pass) & "'"
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("usr_id"))
				setName(rs("usr_name"))
				chkUser = true
			else
				chkUser = false
			end if
			rs.close
			set c = nothing
			
			if chkUser then
				sql = "select a.acs_code from tb_access a inner join tb_usr_acs ac ON a.acs_id = ac.acs_id where ac.usr_id = " & getID
				rs.open sql, conn.conn
				if not (rs.bof and rs.eof) then arrAccess = rs.getRows
				rs.close
			end if
		end function
		
		public sub changeNewPassword(newPass)
			dim sql
			setPassword(newPass)
			
			sql = "select usr_password from tb_user where usr_id = " & getID
			rs.open sql, conn.conn,3,2
			if not (rs.bof and rs.eof) then
				rs("usr_password") = getPassword
				rs.update
			end if
			rs.close
		end sub

		public sub fndUser
			dim sql
			
			sql = "select * from tb_user where usr_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setName(rs("usr_name"))
				setLogin(rs("usr_login"))
				setStatus(rs("usr_status"))
			end if
			rs.close
		end sub
		
		public function ExistsLogin(x)
			dim sql
			ExistsLogin = false
			
			sql = "select 1 from tb_user where usr_login = '" & aspas(x) & "' and usr_id <> " & getId
			rs.open sql,conn.conn
			if not (rs.bof and rs.eof) then
				ExistsLogin = true
			end if
			rs.close
		end function
		
		public sub MngUserAccess(acs)
			dim sql
			
			if getID <> -1 then
				sql = "select * from tb_usr_acs where usr_id = " & getID & " and acs_id = " & acs
				rs.open sql, conn.conn,3,2
				if rs.bof and rs.eof then
					rs.addnew
					rs("usr_id") = getID
					rs("acs_id") = acs
					rs.update
				end if
				rs.close
			end if
		end sub
		
		public sub delUser()
			dim sql
			
			sql = "delete from tb_usr_acs where usr_id = " & getID
			conn.conn.execute(sql)
			sql = "delete from tb_user where usr_id = " & getID
			conn.conn.execute(sql)		
		end sub
		
		public sub delUserAccess(acs)
			dim sql
			
			if getID <> -1 then
				sql = "delete from tb_usr_acs where usr_id = " & getID & " and acs_id = " & acs
				conn.conn.execute(sql)
			end if
		end sub

		public function GetComboAccess(x)
			dim sql, html
			if x <> "" and isNumeric(x) then x = cInt(x)
			
			sql = "select acs_id, acs_section from tb_access order by acs_section"
			rs.open sql, conn.conn
			while not rs.eof
				html = html & "<option value='" & rs("acs_id") & "'"
				if x = rs("acs_id") then html = html & " selected"
				html = html & ">" & rs("acs_section") & "</option>"
			rs.movenext
			wend
			rs.close
			
			GetComboAccess = html
		end function
	end class
%>