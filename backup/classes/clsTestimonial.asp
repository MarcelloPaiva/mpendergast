<%
	class clsTestimonial
		private conn, rs, intID, intName, intDesc, intImage, intStatus
		
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
		
		public sub setDesc(x)
			intDesc = x
		end sub
		
		public sub setImage(x)
			intImage = x
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
		
		public function getDesc()
			getDesc = intDesc
		end function
		
		public function getImage()
			getImage = intImage
		end function
		
		public function getStatus()
			getStatus = intStatus
		end function

		public function MngTestimonial()
			dim sql
			
			sql = "select * from tb_testimonial where tes_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("tes_name") = getName
			rs("tes_desc") = getDesc
			if getImage <> "" then rs("tes_image") = getImage
			rs("tes_status") = getStatus
			rs.update
			rs.close
		end function
		
		public sub fndTestimonial
			dim sql
		
			sql = "select * from tb_testimonial where tes_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("tes_id"))
				setName(rs("tes_name"))
				setDesc(rs("tes_desc"))
				setImage(rs("tes_image"))
				setStatus(rs("tes_status"))
			end if
			rs.close
		end sub
		
		public sub delTestimonial()
			dim sql, obj
			
			fndTestimonial()
			
			if getImage <> "" then
				set obj = new clsDelUploadedFile
				obj.delFile(Application("intUploadPath") & "\" & getImage)
				set obj = nothing
			end if			

			sql = "delete from tb_testimonial where tes_id = " & getID
			conn.conn.execute(sql)
		end sub
	end class
%>