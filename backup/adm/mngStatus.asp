<!--#include file="../globalFunctions/noCache.asp" -->
<!--#include file="../globalFunctions/functions.asp" -->
<!--#include file="../globalFunctions/api.asp" -->
<%
	if not validaNumero(getQuery("id")) then response.Redirect("listNews.asp")
	
	dim id, intStatus, page, table, column, sql, conn
	id = getQuery("id")
	intStatus = getQuery("status")
	table = getQuery("section")
	page = getQuery("url")
	
	select case table
		case "sec_text"
			column = "st"
		case "testimonial"
			column = "tes"
		case "user"
			column = "usr"
		case "aditional_information"
			column = "ai"
		case "useful_link"
			column = "ul"
		case "property"
			column = "pro"
		case "property_image"
			column = "pi"
		case "news"
			column = "new"
		case "sample_ad"
			column = "sa"
		case else
			response.Redirect("listNews.asp")
	end select
	
	if intStatus = "1" then
		intStatus = 1
	else
		intStatus = 0
	end if
	
	sql = "update tb_" & table & " set " & column & "_status = " & intStatus & " where " & column & "_id = " & id
	'response.Write(sql)
	'response.End()
	
	set conn = new clsConnection
	conn.conn.execute(sql)
	set conn = nothing
	
	response.Redirect(page)
%>