<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Templates/castlesclientcnekt.asp" -->
<%
ContactName = Request.Form("ContactName")
EmailAddress = Request.Form("EmailAddress")
TelNumber = Request.Form("TelNumber")
Message = Request.Form("Message")
IPAddress = Request.ServerVariables("REMOTE_ADDR")

Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From = EmailAddress
objCDOMail.Cc = ""
objCDOMail.Bcc = "toffling@dreamingcode.com"
objCDOMail.Subject = "Castles contact email from " & ContactName
objCDOMail.To = "info@castlesmag.com"
objCDOMail.Body =  Message
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