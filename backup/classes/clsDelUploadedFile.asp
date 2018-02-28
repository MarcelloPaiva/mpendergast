<%
	class clsDelUploadedFile
		Sub DelFile(path)
			Dim fs
			Set fs = Server.CreateObject("Scripting.FileSystemObject")
			If (fs.FileExists(path)) Then
				fs.DeleteFile(path)
			End If
			Set fs = Nothing
		End Sub
	end class
%>