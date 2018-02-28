<%@ LANGUAGE="VBSCRIPT" %>

<%
	response.ContentType = "text/xml"
	
	' Create an ASP XML parser object
	set xml = Server.CreateObject("AspXml.AspXml")

    ' This is going to be a collection of contacts.
    xml.Tag = "companies"

    ' Create a child node to hold a contact record.
    ' Our "xml" object will now reference the new node.
    set xml = xml.NewChild("company", "")

    ' Create a "company" record.  NewChild2 does not return the child node that is created.
    xml.NewChild2 "name", "Chilkat Software, Inc."
    xml.NewChild2 "address", "1719 E Forest Ave"
    xml.NewChild2 "city", "Wheaton"
    xml.NewChild2 "state", "IL"
    xml.NewChild2 "zip", "60187"
    xml.NewChild2 "website", "http:'www.chilkatsoft.com"
    xml.NewChild2 "phone", "312-953-3949"

    ' Move up the tree and create another contact record.
    ' GetParent2 updates our internal reference to the parent.
    xml.GetParent2()

    set xml = xml.NewChild("company", "")
    xml.NewChild2 "name", "Microsoft Corporation"
    xml.NewChild2 "address", "One Microsoft Way"
    xml.NewChild2 "city", "Redmond"
    xml.NewChild2 "state", "WA"
    xml.NewChild2 "zip", "98052"
    xml.NewChild2 "website", "http:'www.microsoft.com"
    xml.NewChild2 "phone", "425-882-8080"

    xml.GetParent2()
    set xml = xml.NewChild("company", "")
    xml.NewChild2 "name", "Intel Corporation"
    xml.NewChild2 "address", "2200 Mission College Blvd."
    xml.NewChild2 "city", "Santa Clara"
    xml.NewChild2 "state", "CA"
    xml.NewChild2 "zip", "95052"
    xml.NewChild2 "website", "http:'www.intel.com"
    xml.NewChild2 "phone", "408-765-8080"

    xml.GetParent2()
    set xml = xml.NewChild("company", "")
    xml.NewChild2 "name", "Adobe Systems Incorporated"
    xml.NewChild2 "address", "345 Park Avenue"
    xml.NewChild2 "city", "San Jose"
    xml.NewChild2 "state", "CA"
    xml.NewChild2 "zip", "95110"
    xml.NewChild2 "website", "http:'www.adobe.com"
    xml.NewChild2 "phone", "408-536-6000"

    ' We can always get back to the root by calling GetRoot (or GetRoot2)
    xml.GetRoot2()

    ' Add the tag that specifies our character encoding.
    xml.Encoding = "iso-8859-1"

	response.write xml.GetXml()
%>


