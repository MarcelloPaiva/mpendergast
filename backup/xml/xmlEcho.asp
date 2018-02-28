<%@ language=javascript %>
<%    
Response.Expires = -1000;    
var doc = Server.CreateObject("Msxml2.DOMDocument");    
doc.load(Request);    
Response.ContentType = "text/xml";    
doc.save(Response);
%>
