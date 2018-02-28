<%
	class clsPropertyFeature
		private rs, conn, intPropertyID, intFeatureID, intDesc, intStatus
		
		private sub class_initialize
			intPropertyID = -1
			intFeatureID = -1
			set rs = server.CreateObject("adodb.recordset")
			set conn = new clsConnection
		end sub
		
		private sub class_terminate
			if isObject(rs) then set rs = nothing
			if isObject(conn) then set conn = nothing
		end sub
		
		public sub setPropertyID(x)
			if validanumero(x) then intPropertyID = x
		end sub
		
		public sub setFeatureID(x)
			if validanumero(x) then intFeatureID = x
		end sub
		
		public sub setDesc(x)
			intDesc = x
		end sub

		public sub setStatus(x)
			intStatus = x
		end sub
		
		public function getPropertyID()
			getPropertyID = intPropertyID
		end function
		
		public function getFeatureID()
			getFeatureID = intFeatureID
		end function

		public function getDesc()
			getDesc = intDesc
		end function

		public function getStatus()
			getStatus = intStatus
		end function

		public function MngPropertyFeature()
			dim sql
			
			if getPropertyID = -1 or getFeatureID = -1 then
				MngPropertyFeature = "Invalid Property ID or Feature ID.<br>"
				exit function
			end if
			
			sql = "select * from tb_pro_pf where pro_id = " & getPropertyID & " and pf_id = " & getFeatureID
			rs.open sql, conn.conn, 3,2
			if rs.bof and rs.eof then
				rs.addnew
				rs("pro_id") = getPropertyID
				rs("pf_id") = getFeatureID
				rs("pp_desc") = getDesc
				rs("pp_status") = getStatus
			else
				rs("pp_desc") = getDesc
				rs("pp_status") = getStatus
			end if
			rs.update
			rs.close
		end function
		
		public sub fndPropertyFeature
			dim sql
			
			sql = "select * from tb_pro_pf where pro_id = " & getPropertyID & " and pf_id = " & getFeatureID
			rs.open sql, conn.conn
			if not (rs.bof and rs.eof) then
				setDesc(rs("pp_desc"))
				setStatus(rs("pp_status"))
			end if
			rs.close
		end sub
		
		public sub delPropertyFeature()
			dim sql
			
			sql = "delete from tb_pro_pf where pro_id = " & getPropertyID & " and pf_id = " & getFeatureID
			conn.conn.execute(sql)
		end sub
	end class
%>