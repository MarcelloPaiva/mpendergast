<%
	class clsProperty
		private conn, rs, intID, intCityID, intPropertyTypeID, intClientID, intName, intAddress, intNumber, intDesc
		private intPrice, intTxt1, intTxt2, intImg1, intImg2, intStatus, intSellDate, intVtourURL, intVtourDesc
		private intFeatureID, intFeatureDesc
		
		private sub class_initialize
			intID = -1
			intFeatureID = -1
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
		
		public sub setCityID(x)
			intCityID = x
		end sub
		
		public sub setPropertyTypeID(x)
			intPropertyTypeID = x
		end sub
		
		public sub setClientID(x)
			intClientID = x
		end sub

		public sub setName(x)
			intName = x
		end sub
		
		public sub setAddress(x)
			intAddress = x
		end sub

		public sub setNumber(x)
			intNumber = x
		end sub

		public sub setDesc(x)
			intDesc = x
		end sub

		public sub setPrice(x)
			intPrice = x
		end sub

		public sub setTxt1(x)
			intTxt1 = x
		end sub

		public sub setTxt2(x)
			intTxt2 = x
		end sub

		public sub setImg1(x)
			intImg1 = x
		end sub
		
		public sub setImg2(x)
			intImg2 = x
		end sub
		
		public sub setStatus(x)
			intStatus = x
		end sub

		public sub setSellDate(x)
			intSellDate = x
		end sub
		
		public sub setFeatureID(x)
			intFeatureID = x
		end sub
		
		public sub setFeatureDesc(x)
			intFeatureDesc = x
		end sub

		public sub setVtourURL(x)
			intVtourURL = x
		end sub
		
		public sub setVtourDesc(x)
			intVtourDesc = x
		end sub
		
		public function getID()
			getID = intID
		end function
		
		public function getCityID()
			getCityID = intCityID
		end function

		public function getPropertyTypeID()
			getPropertyTypeID = intPropertyTypeID
		end function

		public function getClientID()
			getClientID = intClientID
		end function

		public function getName()
			getName = intName
		end function
		
		public function getAddress()
			getAddress = intAddress
		end function

		public function getNumber()
			getNumber = intNumber
		end function

		public function getDesc()
			getDesc = intDesc
		end function
		
		public function getPrice()
			getPrice = intPrice
		end function

		public function getTxt1()
			getTxt1 = intTxt1
		end function

		public function getTxt2()
			getTxt2 = intTxt2
		end function
		
		public function getImg1()
			getImg1 = intImg1
		end function

		public function getImg2()
			getImg2 = intImg2
		end function

		public function getStatus()
			getStatus = intStatus
		end function
		
		public function getSellDate()
			getSellDate = intSellDate
		end function

		public function getFeatureID()
			getFeatureID = intFeatureID
		end function

		public function getFeatureDesc()
			getFeatureDesc = intFeatureDesc
		end function

		public function getVtourURL()
			getVtourURL = intVtourURL
		end function
		
		public function getVtourDesc()
			getVtourDesc = intVtourDesc
		end function
		
		public function MngProperty()
			dim sql
			
			sql = "select * from tb_property where pro_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("cit_id") = getCityID
			rs("pt_id") = getPropertyTypeID
			rs("cli_id") = getClientID
			rs("pro_name") = getName
			rs("pro_address") = getAddress
			rs("pro_number") = getNumber
			rs("pro_desc") = getDesc
			rs("pro_price") = getPrice
			rs("pro_txt1") = getTxt1
			rs("pro_txt2") = getTxt2
			if getImg1 <> "" then rs("pro_img1") = getImg1
			if getImg2 <> "" then rs("pro_img2") = getImg2
			rs("pro_status") = getStatus
			if getSellDate <> "" then
				rs("pro_selldate") = year(getSellDate) & "/" & month(getSellDate) & "/" & day(getSellDate)
			else
				rs("pro_sellDate") = null
			end if
			rs("pro_vtour_url") = getVtourURL
			rs("pro_vtour_desc") = getVtourDesc
			rs.update
			rs.close
		end function
		
		public sub fndProperty()
			dim sql
			
			sql = "select * from tb_property where pro_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setCityID(rs("cit_id"))
				setPropertyTypeID(rs("pt_id"))
				setClientID(rs("cli_id"))
				setName(rs("pro_name"))
				setAddress(rs("pro_address"))
				setNumber(rs("pro_number"))
				setDesc(rs("pro_desc"))
				setPrice(rs("pro_price"))
				setTxt1(rs("pro_txt1"))
				setTxt2(rs("pro_txt2"))
				setImg1(rs("pro_img1"))
				setImg2(rs("pro_img2"))
				setStatus(rs("pro_status"))
				setSellDate(rs("pro_sellDate"))
				setVtourURL(rs("pro_vtour_url"))
				setVtourDesc(rs("pro_vtour_desc"))
			end if
			rs.close
		end sub

		public function MngPropertyFeature()
			dim sql
			
			if getID = -1 or getFeatureID = -1 then
				MngFeature = "Invalid client ID or Feature ID.<br>"
				exit function
			end if
			
			sql = "select * from tb_pro_pf where pro_id = " & getID & " and pf_id = " & getFeatureID
			rs.open sql, conn.conn, 3,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("pf_id") = getFeatureID
			rs("pp_desc") = getFeatureDesc
			rs.update
			rs.close
		end function
		
		public sub delProperty()
			dim sql, obj, return
			return = true
			
			sql = "select 1 from tb_property_image where pro_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				return = false
			end if
			rs.close
			
			sql = "select 1 from tb_pro_pf where pro_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				return = false
			end if
			rs.close
			
			if return then
				fndProperty()
				
				set obj = new clsDelUploadedFile
				if getImg1 <> "" then
					obj.delFile(Application("intUploadPath") & "\" & getImg1)
				end if

				if getImg2 <> "" then
					obj.delFile(Application("intUploadPath") & "\" & getImg2)
				end if
				set obj = nothing
				
				sql = "delete from tb_property where pro_id = " & getID
				conn.conn.execute(sql)
			end if
		end sub
	end class
%>