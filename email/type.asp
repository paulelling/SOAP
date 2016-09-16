<%@ Language="JavaScript" %>

<HTML>
<HEAD>
<TITLE>SOAP/XML e-mail service</TITLE>
<%
function calcService()
{
  var doc, stylesheet;
	  
  doc = Server.CreateObject("microsoft.xmldom");
  stylesheet = Server.CreateObject("microsoft.xmldom");
      
  doc.async = false;
  stylesheet.async = false;
      
  doc.load("C:\\ws\\email\\type.xml");
  stylesheet.load("C:\\ws\\email\\type.xsl");
	  
  if (doc.parseError == 0 && stylesheet.parseError == 0)
    return doc.transformNode(stylesheet);
  else
    return "<HTML><BODY>There was a problem loading documents.</BODY></HTML>";
}

Response.Write(calcService());
%>


</HEAD>
<BODY>

</BODY>
</HTML>
