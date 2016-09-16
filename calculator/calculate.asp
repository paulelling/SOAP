<%@ Language=VBScript %>
<html>
<head>
  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
</head>
</html>

 

<%
'SoapServer holds the response url
'xmldom and xmlhttp hold objects to allow XML and SOAP pages to interact
'const strResponseUrl = "http://galt/soap/xmlMortgage.asp"
const strResponseUrl = "http://galt/ws/calculator/mortgage.asp"
set xmldom = server.CreateObject("Microsoft.XMLDOM")
set xmlhttp = server.CreateObject("Microsoft.XMLHTTP")

'connect and send SOAP request to server.asp for a response
xmlhttp.open "POST", strResponseUrl, false
xmlhttp.send(XmlTest)

function XmlTest()
	dim rate, term, amount
	rate = Request.Form("strRate")
	term = Request.Form("strTerm")
	amount = Request.Form("strAmount")
	
	XmlTest = XmlTest & "<SOAP:Envelope xmlns:SOAP='urn:schemas-xmlsoap-org:soap.v1'>"
	XmlTest = XmlTest & "<SOAP:Body>"
	XmlTest = XmlTest & "  <MORTGAGE>"
	XmlTest = XmlTest & "		<RATE>" & rate & "</RATE>"
	XmlTest = XmlTest & "		<TERM>" & term & "</TERM>"
	XmlTest = XmlTest & "		<AMOUNT>" & amount & "</AMOUNT>"
	XmlTest = XmlTest & "  </MORTGAGE>"
	XmlTest = XmlTest & "</SOAP:Body>"
	XmlTest = XmlTest & "</SOAP:Envelope>"
End function

'get response and store in a variable for display
	Set answer = xmlhttp.responseXML
	dim theAnswer
	theAnswer = answer.text

'Stop connection objects
set xmlhttp = nothing
set xmldom = nothing

%>

<head>
  <title>SOAP/XML mortgage service</title>
</head>

<center>
<h3>Mortgage Calculator</h3>
<form name='calc' action='calculate.asp' method='post'>
  <br>Interest rate: <input type='text' name='strRate'>
  <br>Term (years): <input type='text' name='strTerm'>
  <br>Amount: <input type='text' name='strAmount'>
  <p><input type='submit' value='Calculate' name='submit'>
</form>
</center>

<p>
<center>

  <%=theAnswer%>

</center>