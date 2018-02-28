<%
	class clsClient
		private conn, rs, intID, intCityID, intName, intEmail, intAddress, intNumber
		private intZipCode, intAreaCode, intPhone, intJoinDate
		
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
		
		public sub setCityID(x)
			intCityID = x
		end sub
		
		public sub setName(x)
			intName = x
		end sub

		public sub setEmail(x)
			intEmail = x
		end sub
		
		public sub setAddress(x)
			intAddress = x
		end sub
		
		public sub setNumber(x)
			intNumber = x
		end sub

		public sub setZipCode(x)
			intZipCode = x
		end sub

		public sub setAreaCode(x)
			intAreaCode = x
		end sub

		public sub setPhone(x)
			intPhone = x
		end sub
		
		public function getID()
			getID = intID
		end function
		
		public function getCityID()
			getCityID = intCityID
		end function
		
		public function getName()
			getName = intName
		end function

		public function getEmail()
			getEmail = intEmail
		end function
		
		public function getAddress()
			getAddress = intAddress
		end function

		public function getNumber()
			getNumber = intNumber
		end function

		public function getZipCode()
			getZipCode = intZipCode
		end function

		public function getAreaCode()
			getAreaCode = intAreaCode
		end function

		public function getPhone()
			getPhone = intPhone
		end function

		public function getJoinDate()
			getJoinDate = intJoinDate
		end function

		public function MngClient()
			dim sql
			
			sql = "select * from tb_client where cli_id = " & getID
			rs.open sql, conn.conn, 3 ,2
			if rs.bof and rs.eof then
				rs.addnew
			end if
			rs("cit_id") = getCityID
			rs("cli_name") = getName
			rs("cli_email") = getEmail
			rs("cli_address") = getAddress
			rs("cli_number") = getNumber
			rs("cli_zipCode") = getZipCode
			rs("cli_areaCode") = getAreaCode
			rs("cli_phone") = getPhone
			rs.update
			rs.close
		end function
		
		public sub fndClient()
			dim sql
			
			sql = "select * from tb_client where cli_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setCityID(rs("cit_id"))
				setName(rs("cli_name"))
				setEmail(rs("cli_email"))
				setAddress(rs("cli_address"))
				setNumber(rs("cli_number"))
				setZipCode(rs("cli_zipCode"))
				setAreaCode(rs("cli_areaCode"))
				setPhone(rs("cli_Phone"))
				intJoinDate = rs("cli_joinDate")
			end if
			rs.close
		end sub
		
		public function delClient()
			dim sql, return
			return = true
			
			sql = "select 1 from tb_property where cli_id = " & getID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				return = false
			end if
			rs.close
			
			if return then
				sql = "delete from tb_client where cli_id = " & getID
				conn.conn.execute(sql)
			end if
			
			delClient = return
		end function
	end class
%>