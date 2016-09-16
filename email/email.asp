<%@ Language="JavaScript" %>

<%

Response.ContentType = "text/xml";

var doc = Server.CreateObject("Microsoft.XMLDOM");
doc.async = false;
doc.load(Request);

var strFrom = doc.documentElement.childNodes.item(0).selectSingleNode("EMAIL/FROM").text;
var strTo = doc.documentElement.childNodes.item(0).selectSingleNode("EMAIL/TO").text;
var strSubject = doc.documentElement.childNodes.item(0).selectSingleNode("EMAIL/SUBJECT").text;
var strMessage = doc.documentElement.childNodes.item(0).selectSingleNode("EMAIL/MESSAGE").text;

var email = Server.CreateObject("CDONTS.NewMail");
email.From = strFrom + " (Paul Elling)";
email.To = strTo + " (Paul Elling)";
email.Subject = strSubject;
email.Body = strMessage;
email.Send();
email.Nothing;

var returnXML = new String("");

returnXML =  "<SOAP:Envelope xmlns:SOAP='urn:schemas-xmlsoap-org:soap.v1'>";
returnXML += "  <SOAP:Body>";
returnXML += "    <RESPONSE>Your message has been sent.</RESPONSE>";
returnXML += "  </SOAP:Body>";
returnXML += "</SOAP:Envelope>";

Response.Write(returnXML);

%>