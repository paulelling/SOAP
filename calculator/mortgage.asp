
<%
'Allow the SOAP request in XML
Response.ContentType = "text/xml"

'xmldom holds objects to allow XML and SOAP pages to interact
set loanReq = server.CreateObject("Microsoft.XMLDOM")

'The request data
loanReq.load(Request)
set theRate = loanReq.documentElement.childNodes.item(0).selectSingleNode("MORTGAGE/RATE")
set theTerm = loanReq.documentElement.childNodes.item(0).selectSingleNode("MORTGAGE/TERM")
set theAmount = loanReq.documentElement.childNodes.item(0).selectSingleNode("MORTGAGE/AMOUNT")
dim strRate, strTerm, strAmount, first, second, numerator, denominator, annual, monthly
strRate = theRate.text
strTerm = theTerm.text
strAmount = theAmount.text

'Annual loan formula
'a = (P(1 + r)^n * r)  / ((1 + r)^n - 1)

'Calculation
if strRate > 1.0 then
	strRate = strRate / 100.0
end if

first = strRate + 1
second = first ^ strTerm
numerator = (strAmount * second) * strRate
denominator = second - 1
annual = numerator / denominator
monthly = annual / 12

'sum = strRate + strTerm + strAmount

'SOAP response to request from client.asp
returnXML = returnXML & "<SOAP:Envelope xmlns:SOAP='urn:schemas-xmlsoap-org:soap.v1'>"
returnXML = returnXML & "<SOAP:Body>"
returnXML = returnXML & "<RESPONSE>"
returnXML = returnXML & monthly
returnXML = returnXML & "</RESPONSE>"
returnXML = returnXML & "</SOAP:Body>"
returnXML = returnXML & "</SOAP:Envelope>"

'Output the response contents on client.asp
Response.Write(returnXML)

'Stop connection objects
set xmldom = nothing

%>





