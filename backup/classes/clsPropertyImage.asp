<%
	class clsPropertyImage
		private conn, rs, intID, intPropertyID, intTitle, intDesc, intImage, intMain, intAerial, intStatus
		
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
		
		public sub setPropertyID(x)
			intPropertyID = x
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
		
		public sub setMain(x)
			intMain = x
		end sub

		public sub setAerial(x)
			intAerial = x
		end sub

		public sub setStatus(x)
			intStatus = x
		end sub
		
		public function getID()
			getID = intID
		end function

		public function getPropertyID()
			getPropertyID = intPropertyID
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
		
		public function getMain()
			getMain = intMain
		end function

		public function getAerial()
			getAerial = intAerial
		end function

		public function getStatus()
			getStatus = intStatus
		end function

		public function MngPropertyImage()
			dim sql
			
			sql = "select * from tb_property_image where pi_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("pro_id") = getPropertyID
			rs("pi_title") = getTitle
			rs("pi_desc") = getDesc
			rs("pi_image") = getImage
			rs("pi_main") = getMain
			rs("pi_aerial") = getAerial
			rs("pi_status") = getStatus
			rs.update
			rs.close
		end function
		
		public sub fndPropertyImage
			dim sql
			
			sql = "select * from tb_property_image where pi_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setPropertyID(rs("pro_id"))
				setTitle(rs("pi_title"))
				setDesc(rs("pi_desc"))
				setImage(rs("pi_image"))
				setMain(rs("pi_main"))
				setAerial(rs("pi_aerial"))
				setStatus(rs("pi_status"))
			end if
			rs.close
		end sub
		
		public sub delPropertyImage()
			dim sql, obj, img
			
			fndPropertyImage()
			if getImage <> "" then
				set obj = new clsDelUploadedFile
				obj.delFile(Application("intUploadPath") & "\" & getImage)
				set obj = nothing
			end if
		
			sql = "delete from tb_property_image where pi_id = " & getID
			conn.conn.execute(sql)
		end sub
	end class
%>