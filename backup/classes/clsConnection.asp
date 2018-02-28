<%
	class clsConnection
		private strConn
		
		private sub class_initialize
		strConn = "Provider=Microsoft.Ace.OLEDB.12.0; Data Source=//nawinfs02/home/users/web/b713/rh.mpendergast\db\db_pendergast.mdb"
'		strConn = "Provider=Microsoft.Ace.OLEDB.12.0; Data Source=" & Server.MapPath("db/db_pendergast.mdb") & ";"
		end sub
		
		private sub class_terminate
			conn.close
		end sub
		
		public function conn()
			set Conn = server.CreateObject("adodb.connection")
			conn.open(strConn)
		end function
	end class
%>
