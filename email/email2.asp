<%@ Language = "JavaScript" %>

<%

var email = Server.CreateObject("CDONTS.NewMail");
email.From = "paulelling@yahoo.com";
email.To = "pelling@intergate.com";
email.Subject = "howdy";
email.Body = "yo";
email.Send();
email = null;

%>