<%
	class clsCity
		private conn, rs, intID, intName, intStateID
		
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
		
		public sub setStateID(x)
			intStateID = x
		end sub
		
		public sub setName(x)
			intName = x
		end sub

		public function getID()
			getID = intID
		end function
		
		public function getStateID()
			getStateID = intStateID
		end function
		
		public function getName()
			getName = intName
		end function

		public function MngCity()
			dim sql
			
			sql = "select * from tb_city where cit_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("stt_id") = getStateID
			rs("cit_name") = getName
			rs.update
			rs.close
		end function
		
		public sub fndCity()
			dim sql
			
			sql = "select * from tb_city where cit_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("cit_id"))
				setStateID(rs("stt_id"))
				setName(rs("cit_name"))
			end if
			rs.close
		end sub
		
		public function delCity()
			dim sql, return
			return = true
			
			sql = "select 1 from tb_client where cit_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				return = false
			end if
			rs.close
			
			sql = "select 1 from tb_property where cit_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				return = false
			end if
			rs.close

			if return then
				sql = "delete from tb_city where cit_id = " & getID
				conn.conn.execute(sql)
			end if
			
			delCity = return
		end function

		public function GetComboCity(x)
			dim sql, html
			if x <> "" and isNumeric(x) then x = cInt(x)
			
			sql = "select s.stt_abrev, c.cit_id, c.cit_name from tb_city c" &_
				" inner join tb_state s on s.stt_id = c.stt_id" &_
				" order by s.stt_abrev, c.cit_name"
			rs.open sql, conn.conn
			while not rs.eof
				html = html & "<option value='" & rs("cit_id") & "'"
				if x = rs("cit_id") then html = html & " selected"
				html = html & ">" & rs("stt_abrev") & " - " & rs("cit_name") & "</option>"
			rs.movenext
			wend
			rs.close
			
			GetComboCity = html
		end function
	end class
%>