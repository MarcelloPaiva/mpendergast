<%
	class clsSecText
		private conn, rs, intID, intSectionID, intTitle, intDesc, intImage, intFooter, intStatus, intURL
		
		private sub class_initialize
			intID = -1
			intSectionID = -1
			set rs = server.CreateObject("adodb.recordset")
			set conn = new clsConnection
		end sub
		
		private sub class_terminate
			if isObject(rs) then set rs = nothing
			if isObject(conn) then set conn = nothing
		end sub
		
		public sub setID(x)
			if validaNumero(x) then intID = x
		end sub

		public sub setSectionID(x)
			if validaNumero(x) then intSectionID = fndSectionID(x)
		end sub
		
		public sub setTitle(x)
			intTitle = x
		end sub
		
		public sub setDesc(x)
			intDesc = x
		end sub
		
		public sub setImage(x)
			intImage = x
		end sub
		
		public sub setFooter(x)
			intFooter = x
		end sub

		public sub setStatus(x)
			intStatus = x
		end sub
		
		public sub setURL(x)
			intURL = x
		end sub
		
		public function getID()
			getID = intID
		end function
		
		public function getSectionID()
			dim sql 
			
			sql = "select sec_code from tb_section where sec_id = " & intSectionID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				getSectionID = cint(replace(rs("sec_code"), "SE", ""))
			end if
			rs.close
		end function

		public function getTitle()
			getTitle = intTitle
		end function
		
		public function getDesc()
			getDesc = intDesc
		end function
		
		public function getImage()
			getImage = intImage
		end function
		
		public function getFooter()
			getFooter = intFooter
		end function

		public function getStatus()
			getStatus = intStatus
		end function

		public function getURL()
			getURL = intURL
		end function

		public function MngSecText()
			dim sql
			
			if intSectionID = -1 then 
				mngSecText = "Failure when post data.<br>"
				exit function
			end if
			
			sql = "select * from tb_sec_text where st_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("sec_id") = intSectionID
			rs("st_title") = getTitle
			rs("st_desc") = getDesc
			if getImage <> "" then rs("st_image") = getImage
			rs("st_footer") = getFooter
			rs("st_status") = getStatus
			rs("st_url") = getURL
			rs.update
			rs.close
		end function
		
		public sub fndSecText()
			dim sql
			
			sql = "select * from tb_sec_text where st_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("st_id"))
				intSectionID = rs("sec_id")
				setTitle(rs("st_title"))
				setDesc(rs("st_desc"))
				setImage(rs("st_image"))
				setFooter(rs("st_footer"))
				setStatus(rs("st_status"))
				setURL(rs("st_url"))
			end if
			rs.close
		end sub

		public function fndSectionName(x)
			dim sql, section
			section = "SE" & right("000" & x,3)
			
			sql = "select sec_name from tb_section where sec_code = '" & section & "'"
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				fndSectionName = rs("sec_name")
			end if
			rs.close
		end function

		public function fndSectionID(x)
			dim sql, section
			section = "SE" & right("000" & x,3)
			
			sql = "select sec_id from tb_section where sec_code = '" & section & "'"
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				fndSectionID = rs("sec_id")
			end if
			rs.close
		end function
		
		public sub delSecText()
			dim sql, obj
			
			fndSecText()
			
			if getImage <> "" then
				set obj = new clsDelUploadedFile
				obj.delFile(Application("intUploadPath") & "\" & getImage)
				set obj = nothing
			end if
			
			sql = "delete from tb_sec_text where st_id = " & getID
			conn.conn.execute(sql)
		end sub
	end class
%>