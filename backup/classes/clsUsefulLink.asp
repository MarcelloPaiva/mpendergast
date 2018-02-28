<%
	class clsUsefulLink
		private conn, rs, intID, intTitle, intDesc, intURL, intStatus, intType
		
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
		
		public sub setType(x)
			intType = x
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

		public function getType()
			getType = intType
		end function
		
		public function MngUsefulLink()
			dim sql
			
			sql = "select * from tb_useful_link where ul_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("ul_title") = getTitle
			rs("ul_desc") = getDesc
			rs("ul_url") = getURL
			rs("ul_status") = getStatus
			rs("ul_type") = getType
			rs.update
			rs.close
		end function
		
		public sub fndUsefulLink
			dim sql
			
			sql = "select * from tb_useful_link where ul_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("ul_id"))
				setTitle(rs("ul_title"))
				setDesc(rs("ul_desc"))
				setURL(rs("ul_url"))
				setStatus(rs("ul_status"))
				setType(rs("ul_type"))
			end if
			rs.close
		end sub
		
		public sub delUsefulLink()
			dim sql
			
			sql = "delete from tb_useful_link where ul_id = " & getID
			conn.conn.execute(sql)
		end sub
	end class
%>