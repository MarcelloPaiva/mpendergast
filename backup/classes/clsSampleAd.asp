<%
	class clsSampleAd
		private conn, rs, intID, intTitle, intDesc, intImage, intDocument, intURL, intStatus
		
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
		
		public sub setImage(x)
			intImage = x
		end sub

		public sub setDocument(x)
			intDocument = x
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
		
		public function getImage()
			getImage = intImage
		end function
		
		public function getDocument()
			getDocument = intDocument
		end function

		public function getURL()
			getURL = intURL
		end function
		
		public function getStatus()
			getStatus = intStatus
		end function

		public function MngSampleAd()
			dim sql
			
			sql = "select * from tb_sample_ad where sa_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("sa_title") = getTitle
			rs("sa_desc") = getDesc
			if getImage <> "" then rs("sa_image") = getImage
			if getDocument <> "" then rs("sa_doc") = getDocument
			rs("sa_url") = getURL
			rs("sa_status") = getStatus
			rs.update
			rs.close
		end function
		
		public sub fndSampleAd
			dim sql
		
			sql = "select * from tb_Sample_ad where sa_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setID(rs("sa_id"))
				setTitle(rs("sa_title"))
				setDesc(rs("sa_desc"))
				setImage(rs("sa_image"))
				setDocument(rs("sa_doc"))
				setURL(rs("sa_url"))
				setStatus(rs("sa_status"))
			end if
			rs.close
		end sub
		
		public sub delSampleAd()
			dim sql, obj
			
			fndSampleAd()
			
			set obj = new clsDelUploadedFile
			if getImage <> "" then
				obj.delFile(Application("intUploadPath") & "\" & getImage)
			end if			

			if getDocument <> "" then
				obj.delFile(Application("intUploadPath") & "\" & getDocument)
			end if			
			set obj = nothing
				
			sql = "delete from tb_Sample_ad where sa_id = " & getID
			conn.conn.execute(sql)
		end sub
	end class
%>