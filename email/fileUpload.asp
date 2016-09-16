<%@ LANGUAGE="VBScript" %>

<%
Option Explicit
Response.Expires = 0

'define variables and COM objects
dim ado_stream
dim xml_dom
dim xml_file1

'create Stream Object
set ado_stream = Server.CreateObject("ADODB.Stream")

'create XMLDOM object and load it from request ASP object
set xml_dom = Server.CreateObject("MSXML2.DOMDocument")
xml_dom.load(request)

'retrieve XML node with binary content
set xml_file1 = xml_dom.selectSingleNode("root/file1")

'open stream object and store XML node content into it
ado_stream.Type = 1  '1 = adTypeBinary
ado_stream.open
ado_stream.Write xml_file1.nodeTypedValue

'save uploaded file
ado_stream.SaveToFile "google2.gif",2  '2 = adSaveCreateOverWrite
ado_stream.close

'destroy COM object
set ado_stream = Nothing
set xml_dom = Nothing

'write message to browser
Response.Write "Upload successful!"
%>