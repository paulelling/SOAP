<%@ LANGUAGE="JavaScript" %>


<%
//xmldom and xmlhttp hold objects to allow XML and SOAP pages to interact
var xmlhttp = Server.CreateObject("Microsoft.XMLHTTP");

var SoapRequest = new String("");
var	theFrom = new String(Request("strFrom"));
var	theTo = new String(Request("strTo"));
var	theSubject = new String(Request("strSubject"));
var	theMessage = new String(Request("strMessage"));
//var theAttachment = new String(Request("strAttachment"));

	SoapRequest =  "<SOAP-ENV:Envelope xmlns:SOAP-ENV='http://schemas.xmlsoap.org/soap/envelope/'>";
	SoapRequest	+= "  <SOAP-ENV:Body>";
	SoapRequest	+= "    <EMAIL>";
	SoapRequest	+= "      <FROM>" + theFrom.valueOf() + "</FROM>";
	SoapRequest	+= "      <TO>" + theTo.valueOf() + "</TO>";
	SoapRequest	+= "      <SUBJECT>" + theSubject.valueOf() + "</SUBJECT>";
	SoapRequest	+= "      <MESSAGE>" + theMessage.valueOf() + "</MESSAGE>";
	//SoapRequest += "      <ATTACHMENT>" + theAttachment.valueof() + "</ATTACHMENT>";
	//SoapRequest += "      <theSignedForm href='http://zorin/practice/Email/google.gif'/>"; 
	SoapRequest	+= "    </EMAIL>";
	SoapRequest	+= "  </SOAP-ENV:Body>";
	SoapRequest	+= "</SOAP-ENV:Envelope>";


//connect and send SOAP request to server.asp for a response
xmlhttp.open("POST", "http://galt/ws/email/email.asp", false);
xmlhttp.send(SoapRequest);

//if the service works, then write content from client.asp and response content from server.asp
//otherwise, display an error message
if (xmlhttp.Status == 200) 
  var answer = xmlhttp.responseXML.text;
else
  Response.Write("There was a problem reading the SOAP response.");

//Stop connection objects
xmlhttp.Nothing;
%>

<HTML>
  <HEAD>
    <TITLE>
      SOAP/XML e-mail service
    </TITLE>
  </HEAD>
  <BODY>
    <CENTER>
      <%Response.Write(answer)%>
    </CENTER>  
  </BODY>
</HTML>
