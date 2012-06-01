<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Templates/CastlesDCdataDriver.asp" --> 
<%
'Persist Search Values
Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
SearchColumn = Request.QueryString("SearchColumn")
SearchValue = Request.QueryString("SearchValue")
redirectTo = Request.QueryString("redirectTo")

'Necessary Requests
SystemLoginID = Request.Cookies("SystemLoginID")
IPAddress = Request.ServerVariables("REMOTE_ADDR")
Active = Request.Form("Active")
HeaderCount = Request.Form("HeaderCount")
SubHeaderCount = Request.Form("SubHeaderCount")
UserEntityID = Request.Form("EntityID")
EntityPrimaryKeyValue = Request.Form("EntityPrimaryKeyValue")
'AdminLanguageID = Request.Form("AdminLanguageID")
NickName = CleanForDrive(Request.Form("NickName"))
UserName = CleanForDrive(Request.Form("UserName"))
Password = CleanForDrive(Request.Form("Password"))
EmailAddress = CleanForDrive(Request.Form("EmailAddress"))
ShowHelpContent = CleanForDrive(Request.Form("ShowHelpContent"))
CreateNotes = CleanForDrive(Request.Form("CreateNotes"))
If Len(CreateNotes) = 0 Then
	CreateNotes = "N"
End If
ViewNotes = CleanForDrive(Request.Form("ViewNotes"))
If Len(ViewNotes) = 0 Then
	ViewNotes = "N"
End If
DeleteNotes = CleanForDrive(Request.Form("DeleteNotes"))
If Len(DeleteNotes) = 0 Then
	DeleteNotes = "N"
End If
FirstName = CleanForDrive(Request.Form("FirstName"))

'Syncronizes Cookies Self Update Of Administrator
If cStr(SystemLoginID) = cStr(SystemLoginID) Then
	Response.Cookies("LanguageID") = AdminLanguageID
	Response.Cookies("ShowHelpContent") = ShowHelpContent
	Response.Cookies("SystemLoginNickName") = NickName
	Response.Cookies("CreateNotes") = CreateNotes
	Response.Cookies("ViewNotes") = ViewNotes
	Response.Cookies("DeleteNotes") = DeleteNotes
End If


DataExceptionsString = ""

'Excludes Access Rights Form Fields From Admin Profile
Count = 1
For i = 1 to HeaderCount
	DataExceptionsString = DataExceptionsString & "<!TopNavigationHeaderID" & Count & "!>"
	Count = Count + 1
Next

Count = 1
For i = 1 to SubHeaderCount
	DataExceptionsString = DataExceptionsString & "<!TopNavigationSubHeaderID" & Count & "!><!TopNavigationHeaderIDForSubHeader" & Count & "!>"
	Count = Count + 1
Next

'Which DCdataDriverType to use
DCdataDriverType = Request.QueryString("DCdataDriverType")

'Broker Entity = 7
'Admin Entity = 1
'Designer Entity = 5
'TeleSales = 4
Select Case UserEntityID
	Case 1 ' Admin
		DBEntity = "Castles_Administrators"
		DBUniqueKey = "AdministratorID"
	Case 4 ' TeleSales
		DBEntity = "Castles_TeleSales"
		DBUniqueKey = "TeleSalesID"
	Case 5 ' Designers
		DBEntity = "Castles_Designers"
		DBUniqueKey = "DesignerID"
	Case 7 ' Brokers
		DBEntity = "Castles_Brokers"
		DBUniqueKey = "BrokerID"
End Select

'Handle different DCdataDriverTypes
Select Case DCdataDriverType
	Case "SQLInsert"
		'...

	Case "SQLUpdate"
		'Checks for duplicate username
		UserName = CleanForDrive(Request.Form("UserName"))
		SystemUserID = CleanForDrive(Request.Form("SystemUserID"))
		DuplicateFound = "N"
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_SystemLogin_CheckFor_DupUserName_OnEdit"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@SystemLoginID", 200, 1,200,SystemLoginID)
			.Parameters.Append .CreateParameter("@UserName", 200, 1,200,UserName)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set DupUserName = .Execute()
		End With
		Set Command1 = Nothing
		
		If Not DupUserName.EOF Then
			DuplicateFound = "Y"
			OldUserName = CleanForDrive(Request.Form("OldUserName"))
			AttemptUserName = CleanForDrive(Request.Form("UserName"))
		End If

		'Updates Administrator Information
		DBObjectDestination = DBEntity
		FileServerDestination = ""
		FileFate = ""
		FormType = "Request.Form"
		EmailDestination = ""
		DataUniqueKey = DBUniqueKey
		DataParentNode = ""
		DataExceptions = "<!SystemLoginID!><!EntityID!><!BrokerID!><!AdministratorID!><!TeleSalesID!><!DesignerID!><!HeaderCount!><!SubHeaderCount!><!NickName!><!UserName!><!Password!><!ShowHelpContent!><!CreateNotes!><!ViewNotes!><!DeleteNotes!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		If DuplicateFound = "N" Then
			DCdataDriverExecutedRedirectURL = "editProfile.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		Else
			DCdataDriverExecutedRedirectURL = "editProfile.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords & "&DuplicateFound=" & DuplicateFound & "&OldUserName=" & Server.URLEncode(UserName) & "&AttemptUserName=" & Server.URLEncode(AttemptUserName)
		End If
		Cnekt = Connect
		
		DCdataDriver DCdataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Tracks Modification of Administrator Account
		EntityID = 1
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,EntityPrimaryKeyValue,EntityModificationTypeID,IPAddress 

		'Updates System Login Username & Password
		Deleted = "N"
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Connect
		SQLStmt = "UPDATE Castles_SystemLogins SET "
		SQLStmt = SQLStmt & "NickName ='" & NickName & "',"
		SQLStmt = SQLStmt & "UserName ='" & UserName & "',"
		SQLStmt = SQLStmt & "Password ='" & Password & "',"
		SQLStmt = SQLStmt & "EmailAddress ='" & EmailAddress & "',"
		SQLStmt = SQLStmt & "LanguageID ='" & LanguageID & "',"
		SQLStmt = SQLStmt & "ShowHelpContent ='" & ShowHelpContent & "',"
		SQLStmt = SQLStmt & "CreateNotes ='" & CreateNotes & "',"
		SQLStmt = SQLStmt & "ViewNotes ='" & ViewNotes & "',"
		SQLStmt = SQLStmt & "Active ='" & Active & "',"
		SQLStmt = SQLStmt & "Deleted ='" & Deleted & "',"
		SQLStmt = SQLStmt & "DeleteNotes ='" & DeleteNotes & "' "
		SQLStmt = SQLStmt & "WHERE SystemLoginID ='" & SystemLoginID & "'"
		Set SQLAction = Conn.Execute(SQLStmt)
		Set SQLAction = Nothing
		Conn.Close

		'Tracks Modification of System Login Info
		EntityID = 6
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,SystemLoginID,EntityModificationTypeID,IPAddress 

		DCdataDriverExecutedRedirectURL = "editProfile.asp?success=Y"
		
		Response.Redirect DCdataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'...
End Select
%>
