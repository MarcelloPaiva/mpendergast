<%
	class clsState
		private conn, rs, intID, intName, intAbreviation
		
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
		
		public sub setAbreviation(x)
			intAbreviation = x
		end sub

		public function getID()
			getID = intID
		end function
		
		public function getName()
			getName = intName
		end function
		
		public function getAbreviation()
			getAbreviation = intAbreviation
		end function

		public function MngState()
			dim sql
			
			sql = "select * from tb_state where stt_id = " & getID
			'response.Write(sql)
			'response.End()
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("stt_name") = getName
			rs("stt_abrev") = getAbreviation
			rs.update
			rs.close
		end function
		
		public sub fndState()
			dim sql
			
			sql = "select * from tb_state where stt_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("stt_id"))
				setAbreviation(rs("stt_abrev"))
				setName(rs("stt_name"))
			end if
			rs.close
		end sub
		
		public function delState()
			dim sql
			
			sql = "select 1 from tb_city where stt_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				delState = false
			else
				sql = "delete from tb_state where stt_id = " & getID
				conn.conn.execute(sql)
				delState = true
			end if
			rs.close
		end function
		
		public function GetComboState(x)
			dim sql, html
			if x <> "" and isNumeric(x) then x = cInt(x) else x = ""
			
			sql = "select stt_id, stt_name, stt_abrev from tb_state order by stt_name"
			rs.open sql, conn.conn
			if x = "" then
				while not rs.eof
					html = html & "<option value='" & rs("stt_id") & "'"
					if rs("stt_abrev") = "MA" then html = html & " selected"
					html = html & ">" & rs("stt_name") & "</option>"
				rs.movenext
				wend
			else
				while not rs.eof
					html = html & "<option value='" & rs("stt_id") & "'"
					if x = rs("stt_id") then html = html & " selected"
					html = html & ">" & rs("stt_name") & "</option>"
				rs.movenext
				wend
			end if
			rs.close
			
			GetComboState = html
		end function
	end class
%>