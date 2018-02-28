<%
	class clsNews
		private conn, rs, intID, intTitle, intDesc, intDate, intReference, intStatus, intEndDate
		
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
		
		public sub setTitle(x)
			intTitle = x
		end sub
		
		public sub setDesc(x)
			intDesc = x
		end sub
		
		public sub setDate(x)
			intDate = x
		end sub
		
		public sub setReference(x)
			intReference = x
		end sub
		
		public sub setStatus(x)
			intStatus = x
		end sub

		public sub setEndDate(x)
			if isDate(x) then intEndDate = x
		end sub

		public function getID()
			getID = intID
		end function
		
		public function getTitle()
			getTitle = intTitle
		end function
		
		public function getDesc()
			getDesc = intDesc
		end function

		public function getDate()
			getDate = intDate
		end function

		public function getReference()
			getReference = intReference
		end function
		
		public function getStatus()
			getStatus = intStatus
		end function
		
		public function getEndDate
			getEndDate = intEndDate
		end function
		
		public function MngNews()
			dim sql
			
			sql = "select * from tb_news where new_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("new_title") = getTitle
			rs("new_desc") = getDesc
			rs("new_reference") = getReference
			rs("new_status") = getStatus
			rs("new_endDate") = getEndDate
			rs.update
			rs.close
		end function
		
		public sub fndNews()
			dim sql
			
			sql = "select * from tb_news where new_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("new_id"))
				setTitle(rs("new_title"))
				setDesc(rs("new_desc"))
				setDate(rs("new_date"))
				setReference(rs("new_reference"))
				setStatus(rs("new_status"))
				setEndDate(rs("new_endDate"))
			end if
			rs.close
		end sub
		
		public function delNews()
			dim sql
			
			sql = "delete from tb_news where new_id = " & getID
			conn.conn.execute(sql)
		end function
	end class
%>