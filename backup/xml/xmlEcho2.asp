<%@ language=javascript %>
<%    
Response.Expires = -1000;    
var doc = Server.CreateObject("AspXml.AspXml");    
doc.LoadRequest();    
Response.ContentType = "text/xml";    
Response.Write(doc.GetXml());

%>
