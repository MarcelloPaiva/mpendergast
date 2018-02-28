<%
	class clsPropertyType
		private conn, rs, intID, intName
		
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

		public function getID()
			getID = intID
		end function
		
		public function getName()
			getName = intName
		end function

		public function MngPropertyType()
			dim sql
			
			sql = "select * from tb_property_type where pt_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("pt_name") = getName
			rs.update
			rs.close
		end function
		
		public sub fndPropertyType
			dim sql
			
			sql = "select * from tb_property_type where pt_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("pt_id"))
				setName(rs("pt_name"))
			end if
			rs.close
		end sub
		
		public function GetComboPropertyType(x)
			dim sql, html
			if x <> "" and isNumeric(x) then x = cInt(x)
			
			sql = "select pt_id, pt_name from tb_property_type order by pt_name"
			rs.open sql, conn.conn
			while not rs.eof
				html = html & "<option value='" & rs("pt_id") & "'"
				if x = rs("pt_id") then html = html & " selected"
				html = html & ">" & rs("pt_name") & "</option>"
			rs.movenext
			wend
			rs.close
			
			GetComboPropertyType = html
		end function

		public function delPropertyType()
			dim sql
			
			sql = "select 1 from tb_property where pt_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				delPropertyType = false
			else
				sql = "delete from tb_Property_Type where pt_id = " & getID
				conn.conn.execute(sql)
				delPropertyType = true
			end if
			rs.close
		end function
	end class
%>