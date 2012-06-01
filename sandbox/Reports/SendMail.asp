<%@ Language=VBScript %>
<html>
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
</head>
<body>
<%if Request.QueryString("Visited")<>"Yes" Then%>
<form action="default.asp?Visited=Yes" name=MailInformation method=post>
	<INPUT type=checkbox value="9 Deadly Mistakes Home Sellers Make" name=doc1 ID="Checkbox1"> 9 Deadly Mistakes Home Sellers Make<BR>
	<INPUT type=checkbox value="Making the Move Easy on the Kids" name=doc2 ID="Checkbox2"> Making the Move Easy on the Kids<BR>
	<INPUT type=checkbox value="Some Different Reasons to Own Your Own Home" name=doc3 ID="Checkbox3"> Some Different Reasons to Own Your Own Home<BR>
	<INPUT type=checkbox value="14 Questions to ask a Realtor" name=doc4 ID="Checkbox4"> 14 Questions to ask a Realtor<BR>
	<INPUT type=checkbox value="5 Powerful Buying Strategies" name=doc5 ID="Checkbox5"> 5 Powerful Buying Strategies<BR>
	<INPUT type=checkbox value="Six Ways To Beat The Stress Of Buying A Home" name=doc6 ID="Checkbox6"> Six Ways To Beat The Stress Of Buying A Home<BR>
	<INPUT type=checkbox value="Things You Should Know about Moving" name=doc7 ID="Checkbox7"> Things You Should Know about Moving<BR>
	<INPUT type=checkbox value="When Selling a Home" name=doc8 ID="Checkbox8"> When Selling a Home<BR>
	<INPUT type=checkbox value="How To Get Top Dollar In Any Market" name=doc9 ID="Checkbox9"> How To Get Top Dollar In Any Market<BR><BR>
	<P>Please take a moment to fill out the following information:<BR>
	<BR>Name:<BR><INPUT name=name ID="Text1">
	<BR>Phone:<BR><INPUT name=phone ID="Text2">
	<BR>Email:<BR><INPUT name=email ID="Text3">
	<BR>* please note that you must fill out all fields<BR>
	<BR><INPUT type=submit value="Get Reports" ID="Submit1" NAME="Submit1">
</form>
<%End If%>
<%If Request.QueryString("Visited")="Yes" Then
 Response.Write("The following Reorts have been sent to your mail.<br>")
 
	for each Fld in request.Form
		FldValue=Request.Form(Fld)
		If FldValue<> "" Then
			If Instr(Fld,"doc")<> 0 Then
							Response.Write (FldValue&"<br>")
							set ObjSendMailToSignIn= CreateObject("CDONTS.NewMail")
							ObjSendMailToSignIn.From="Testing@DreamingCode.com"
							ObjSendMailToSignIn.To= Request.Form("Email")
							ObjSendMailToSignIn.Subject= "Report on "&FldValue
							ObjSendMailToSignIn.BodyFormat = 0
							ObjSendMailToSignIn.MailFormat = 0
							ObjSendMailToSignIn.Body= "The content and the attachments are yet to be added"
							ObjSendMailToSignIn.send
							Set ObjSendMailToSignIn=Nothing
			End If 
		End If
		
	Next
End If	
%>
</body>
</html>
