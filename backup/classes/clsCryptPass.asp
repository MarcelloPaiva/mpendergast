<%
	class clsCryptPass
		public function Crypt(text)
			dim x, numero, n, str
			x = text
			numero = ""
			for n = 1 to len(x)
				str = mid(x, n, 1)
				if IsNumeric(str) then
					numero = numero & hex(str^5)
				else
					numero = numero & hex(asc(str)^3)
				end if
			next	
			numero = right("0000000000" & numero,10)
			crypt = numero
		end function
	end class
%>