<%
	class clsAditionalInfo
		private conn, rs, intID, intTitle, intDesc, intURL, intStatus
		
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
		
		public sub setURL(x)
			intURL = x
		end sub
		
		public sub setStatus(x)
			intStatus = x
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
		
		public function getURL()
			getURL = intURL
		end function
		
		public function getStatus()
			getStatus = intStatus
		end function

		public function MngAditionalInfo()
			dim sql
			
			sql = "select * from tb_aditional_information where ai_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("ai_title") = getTitle
			rs("ai_desc") = getDesc
			rs("ai_url") = getURL
			rs("ai_status") = getStatus
			rs.update
			rs.close
		end function
		
		public sub fndAditionalInfo
			dim sql
			
			sql = "select * from tb_aditional_information where ai_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("ai_id"))
				setTitle(rs("ai_title"))
				setDesc(rs("ai_desc"))
				setURL(rs("ai_url"))
				setStatus(rs("ai_status"))
			end if
			rs.close
		end sub
		
		public sub delAditionalInfo()
			dim sql
			
			sql = "delete from tb_aditional_information where ai_id = " & getID
			conn.conn.execute(sql)
		end sub
	end class
%>