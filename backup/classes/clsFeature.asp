<%
	class clsFeature
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

		public function MngFeature()
			dim sql
			
			sql = "select * from tb_feature where pf_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("pf_name") = getName
			rs.update
			rs.close
		end function
		
		public sub fndFEature
			dim sql
			
			sql = "select * from tb_feature where pf_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setName(rs("pf_name"))
			end if
			rs.close
		end sub
		
		public function delFeature()
			dim sql
			
			sql = "select 1 from tb_pro_pf where pf_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				delFeature = false
			else
				sql = "delete from tb_feature where pf_id = " & getID
				conn.conn.execute(sql)
				delFeature = true
			end if
			rs.close
		end function

		public function GetComboFeature(x)
			dim sql, html
			if x <> "" and isNumeric(x) then x = cInt(x)
			
			sql = "select pf_id, pf_name from tb_feature order by pf_name"
			rs.open sql, conn.conn
			while not rs.eof
				html = html & "<option value='" & rs("pf_id") & "'"
				if x = rs("pf_id") then html = html & " selected"
				html = html & ">" & rs("pf_name") & "</option>"
			rs.movenext
			wend
			rs.close
			
			GetComboFeature = html
		end function
	end class
%>