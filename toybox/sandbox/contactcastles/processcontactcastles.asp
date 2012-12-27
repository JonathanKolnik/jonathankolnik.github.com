<!--METADATA TYPE="typelib"
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"
NAME="CDO for Windows 2000 Library" -->
<!--METADATA TYPE="typelib"
UUID="00000205-0000-0010-8000-00AA006D2EA4"
NAME="ADODB Type Library" -->
<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Templates/castlesclientcnekt.asp" -->
<%
ContactName = Request.Form("ContactName")
EmailAddress = Request.Form("EmailAddress")
TelNumber = Request.Form("TelNumber")
Message = Request.Form("Message")
IPAddress = Request.ServerVariables("REMOTE_ADDR")

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
objCDOMail.From = EmailAddress
objCDOMail.Cc = ""
'objCDOMail.Bcc = "toffling@dreamingcode.com"
objCDOMail.Subject = "Castles contact email from " & ContactName
objCDOMail.To = "info@castlesmag.com"
objCDOMail.TextBody =  Message
objCDOMail.Send
Set objCDOMail = Nothing 

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Connect
SQLStmt = "INSERT INTO Castles_ContactUsLog (ContactName,EmailAddress,TelNumber,Message,ContactIPAddress,ContactDateTime) "
SQLStmt = SQLStmt & "VALUES (" & "'" & CleanForDrive(ContactName) & "'" 
SQLStmt = SQLStmt & "," & "'" & CleanForDrive(EmailAddress) & "'" 
SQLStmt = SQLStmt & "," & "'" & CleanForDrive(TelNumber) & "'" 
SQLStmt = SQLStmt & "," & "'" & CleanForDrive(Message) & "'" 
SQLStmt = SQLStmt & "," & "'" & IPAddress & "'" 
SQLStmt = SQLStmt & "," & "'" & Now & "'" & ")"
Set SQLAction = Conn.Execute(SQLStmt)
Set SQLAction = Nothing
Conn.Close

Response.Redirect "ContactCastlesSuccess.asp?name=" & Server.URLEncode(ContactName)
%>