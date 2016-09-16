<%
If Len(Request.Form("txtEmail")) > 0 then

	Dim objMail
	Set objMail = Server.CreateObject("CDONTS.NewMail")

	objMail.From = "paulelling@yahoo.com (Paul Elling)"
	objMail.Subject = "Email attachment demo"
	objMail.AttachFile Server.MapPath("google.gif")
	objMail.To = Request.Form("txtEmail")
	objMail.Body = "This is a demo on sending an email with an attachment."
	objMail.Send

	Response.write("<i>Mail was Sent</i><p>")

	'You should always do this with CDONTS.
	set objMail = nothing
			
	
End If
%>

  <form method="post" id=form1 name=form1>
    <b>Enter your email address:</b><br>
    <input type="text" name="txtEmail" value="<%=Request.Form("txtEmail")%>">
    <p>
    <input type="submit" value="Send me an Email with an Attachment!" id=submit1 name=submit1>
  </form>
  
