<!--METADATA TYPE="typelib"
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"
NAME="CDO for Windows 2000 Library" -->
<!--METADATA TYPE="typelib"
UUID="00000205-0000-0010-8000-00AA006D2EA4"
NAME="ADODB Type Library" -->
<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/templates/Castlessystemcnektonly.asp" -->
<%
EmailAddress = Request.Form("EmailAddress")
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_ForgottenPassword_Retrieve"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@EmailAddress", 200, 1,200,EmailAddress)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set Retriever = .Execute()
End With
Set Command1 = Nothing
if Retriever.EOF then
	response.redirect "forgotpassword.asp?success=N"
else
	While NOT Retriever.EOF
		UserName = Retriever.Fields.Item("UserName").Value
		Password = Retriever.Fields.Item("Password").Value
	Retriever.MoveNext()
	Wend
end if
'Sends Email to Castles Notifying of Recent Sign Up

Message = "This is a message from CastlesMag.com" & vbcrlf & vbcrlf & "You have requested your username and password be sent to you." & vbcrlf & vbcrlf & "Username: " & UserName & vbcrlf & "Password: " & Password & vbcrlf & vbcrlf & "Sincerely," & vbcrlf & "The Castles Magazine Team"
		
Set objCDOMail = Server.CreateObject("CDO.Message")

Set objConfig = Server.CreateObject("CDO.Configuration")
'Configuration:
objConfig.Fields(cdoSendUsingMethod) = cdoSendUsingPort
objConfig.Fields(cdoSMTPServer)="smtp.1and1.com"
objConfig.Fields(cdoSMTPServerPort)=25
objConfig.Fields(cdoSMTPAuthenticate)=cdoBasic
objConfig.Fields(cdoSendUserName) = "m39707745-1"
objConfig.Fields(cdoSendPassword) = "test123"
'Update configuration
objConfig.Fields.Update
Set objCDOMail.Configuration = objConfig
objCDOMail.From = "info@CastlesMag.com"
objCDOMail.Cc = ""
objCDOMail.Bcc = ""
objCDOMail.Subject = "Castles Management System Log-in Information"
objCDOMail.To = EmailAddress
objCDOMail.TextBody =  Message
objCDOMail.Send
Set objCDOMail = Nothing 

Response.Redirect "forgotpassword.asp?success=Y"
%>