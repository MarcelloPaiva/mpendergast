<%
'Sending a text email 
Set objMessage = Server.CreateObject("CDO.Message") 
objMessage.Subject = "Example CDO Message" 
objMessage.From = "martha@marthapendergastrealestate.com" 
objMessage.To = "jlemosmoreira@gmail.com" 
objMessage.TextBody = "This is some sample message text." 
objMessage.Send 
%>

