<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Templates/castlesdcdatadriver.asp" --> 
<%
'Persist Search Values
Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
SearchColumn = Request.QueryString("SearchColumn")
SearchValue = Request.QueryString("SearchValue")

'Necessary Requests
SystemLoginID = Request.Cookies("SystemLoginID")
IPAddress = Request.ServerVariables("REMOTE_ADDR")
Active = Request.Form("Active")
HeaderCount = Request.Form("HeaderCount")
SubHeaderCount = Request.Form("SubHeaderCount")
EntitySystemLoginID = Request.Form("EntitySystemLoginID")
BrokerID = Request.Form("BrokerID")
EntityLanguageID = Request.Form("EntityLanguageID")
NickName = CleanForDrive(Request.Form("NickName"))
UserName = CleanForDrive(Request.Form("UserName"))
Password = CleanForDrive(Request.Form("Password"))
ShowHelpContent = CleanForDrive(Request.Form("ShowHelpContent"))
FirstName = CleanForDrive(Request.Form("FirstName"))

CreateNotes = CleanForDrive(Request.Form("CreateNotes"))
If Len(CreateNotes) = 0 Then
	CreateNotes = "N"
End If
DeleteNotes = CleanForDrive(Request.Form("DeleteNotes"))
If Len(DeleteNotes) = 0 Then
	DeleteNotes = "N"
End If

'Syncronizes Cookies Self Update Of Broker
If cStr(SystemLoginID) = cStr(EntitySystemLoginID) Then
	Response.Cookies("LanguageID") = EntityLanguageID
	Response.Cookies("ShowHelpContent") = ShowHelpContent
	Response.Cookies("SystemLoginNickName") = NickName
	Response.Cookies("CreateNotes") = CreateNotes
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

'Which DCDataDriverType to use
DCDataDriverType = Request.QueryString("DCDataDriverType")

'Handle different DCDataDriverTypes
Select Case DCDataDriverType
	Case "SQLInsert"
		'Checks for duplicate username
		UserName = CleanForDrive(Request.Form("UserName"))
		DuplicateFound = "N"
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_SystemLogin_CheckFor_DupUserName_OnCreate"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@UserName", 200, 1,200,UserName)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set DupUserName = .Execute()
		End With
		Set Command1 = Nothing
		
		If Not DupUserName.EOF Then
			DuplicateFound = "Y"
			UserName = CleanForDrive(Request.Form("EmailAddress"))
			AttemptUserName = CleanForDrive(Request.Form("UserName"))
		End If

		'Inserts Broker Information
		DBObjectDestination = "Castles_Brokers"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "BrokerID"
		DataParentNode = ""
		DataExceptions = "<!BrokerID!><!EntityLanguageID!><!EntitySystemLoginID!><!HeaderCount!><!SubHeaderCount!><!NickName!><!UserName!><!Password!><!ShowHelpContent!><!CreateNotes!><!ViewNotes!><!DeleteNotes!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		Cnekt = Connect
		
		BrokerID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)
		
		'Tracks Modification of Broker Account
		EntityID = 7
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,BrokerID,EntityModificationTypeID,IPAddress 

		'Creates System Login Username & Password
		SystemLoginTypeID = 3
		EntityID = 7
		Deleted = "N"
		EntityPrimaryKeyValue = BrokerID
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Connect
		SQLStmt = "INSERT INTO Castles_SystemLogins (NickName,UserName,Password,LanguageID,ShowHelpContent,CreateNotes,ViewNotes,DeleteNotes,SystemLoginTypeID,Active,Deleted,EntityID,EntityPrimaryKeyValue) "
		SQLStmt = SQLStmt & "VALUES (" & "'" & NickName & "'" 
		SQLStmt = SQLStmt & "," & "'" & UserName & "'" 
		SQLStmt = SQLStmt & "," & "'" & Password & "'" 
		SQLStmt = SQLStmt & "," & "'" & EntityLanguageID & "'" 
		SQLStmt = SQLStmt & "," & "'" & ShowHelpContent & "'" 
		SQLStmt = SQLStmt & "," & "'" & CreateNotes & "'" 
		SQLStmt = SQLStmt & "," & "'" & ViewNotes & "'" 
		SQLStmt = SQLStmt & "," & "'" & DeleteNotes & "'" 
		SQLStmt = SQLStmt & "," & "'" & SystemLoginTypeID & "'" 
		SQLStmt = SQLStmt & "," & "'" & Active & "'" 
		SQLStmt = SQLStmt & "," & "'" & Deleted & "'" 
		SQLStmt = SQLStmt & "," & "'" & EntityID & "'" 
		SQLStmt = SQLStmt & "," & "'" & EntityPrimaryKeyValue & "'" & ")"
		Set SQLAction = Conn.Execute(SQLStmt)
		Set SQLAction = Nothing
		Conn.Close

		'Pulls Out Most Recent SystemLoginID
		GenerateIDSQLStmt =  "SELECT Max(SystemLoginID) AS SystemLoginID FROM Castles_SystemLogins"
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Cnekt
			.CommandText = "DCDataDriver_SQLInsert_GenerateID"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4) 
			.Parameters.Append .CreateParameter("@GenerateIDSQLStmt", 200, 1,200,GenerateIDSQLStmt)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = True
			Set GenerateID = .Execute()
		End With
		Set Command1 = nothing
		If Not GenerateID.EOF Then
			CreatedSystemLoginID = GenerateID.Fields.Item("SystemLoginID").Value
		End If

		'Tracks Modification of System Login Info
		EntityID = 8
		EntityModificationTypeID = 1
		EntityModificationLogger SystemLoginID,EntityID,CreatedSystemLoginID,EntityModificationTypeID,IPAddress 

		'Attaches SystemLoginID to Broker
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Connect
		SQLStmt = "UPDATE Castles_Brokers SET "
		SQLStmt = SQLStmt & "SystemLoginID ='" & CreatedSystemLoginID & "' "
		SQLStmt = SQLStmt & "WHERE BrokerID ='" & EntityPrimaryKeyValue & "'"
		Set SQLAction = Conn.Execute(SQLStmt)
		Set SQLAction = Nothing
		Conn.Close

		'Tracks Modification of System Login Info
		EntityID = 8
		EntityModificationTypeID = 1
		EntityModificationLogger SystemLoginID,EntityID,CreatedSystemLoginID,EntityModificationTypeID,IPAddress 

		'Updates System Login Access Rights
		Accessable = "Y"

		'Header Access Rights Insert
		Count = 1
		For i = 1 to HeaderCount
			If Len(Request.Form("TopNavigationHeaderID" & Count)) <> 0 Then
				TopNavigationHeaderID = Request.Form("TopNavigationHeaderID" & Count)
				TopNavigationSubHeaderID = 0
				Set Conn = Server.CreateObject("ADODB.Connection")
				Conn.Open Connect
				SQLStmt = "INSERT INTO Castles_SystemLoginAccessRights (SystemLoginID,TopNavigationHeaderID,TopNavigationSubHeaderID,LanguageID,Accessable) "
				SQLStmt = SQLStmt & "VALUES (" & "'" & CreatedSystemLoginID & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationHeaderID & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationSubHeaderID & "'" 
				SQLStmt = SQLStmt & "," & "'" & EntityLanguageID & "'" 
				SQLStmt = SQLStmt & "," & "'" & Accessable & "'" & ")"
				Set SQLAction = Conn.Execute(SQLStmt)
				Set SQLAction = Nothing
				Conn.Close

				'Tracks Insertion of Broker Header Access Rights
				EntityID = 9
				EntityModificationTypeID = 1
				EntityModificationLogger SystemLoginID,EntityID,CreatedSystemLoginID,EntityModificationTypeID,IPAddress 
			End If
			Count = Count + 1
		Next

		'SubHeader Access Rights Insert
		Count = 1
		For i = 1 to SubHeaderCount
			If Len(Request.Form("TopNavigationSubHeaderID" & Count)) <> 0 Then
				TopNavigationHeaderIDForSubHeader = Request.Form("TopNavigationHeaderIDForSubHeader" & Count)
				TopNavigationSubHeaderID = Request.Form("TopNavigationSubHeaderID" & Count)
				Set Conn = Server.CreateObject("ADODB.Connection")
				Conn.Open Connect
				SQLStmt = "INSERT INTO Castles_SystemLoginAccessRights (SystemLoginID,TopNavigationHeaderID,TopNavigationSubHeaderID,LanguageID,Accessable) "
				SQLStmt = SQLStmt & "VALUES (" & "'" & CreatedSystemLoginID & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationHeaderIDForSubHeader & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationSubHeaderID & "'" 
				SQLStmt = SQLStmt & "," & "'" & EntityLanguageID & "'" 
				SQLStmt = SQLStmt & "," & "'" & Accessable & "'" & ")"
				Set SQLAction = Conn.Execute(SQLStmt)
				Set SQLAction = Nothing
				Conn.Close

				'Tracks Insertion of Broker SubHeader Access Rights
				EntityID = 9
				EntityModificationTypeID = 1
				EntityModificationLogger SystemLoginID,EntityID,CreatedSystemLoginID,EntityModificationTypeID,IPAddress 
			End If
			Count = Count + 1
		Next

		If DuplicateFound = "N" Then
			DCDataDriverExecutedRedirectURL = "searchBrokers.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		Else
			DCDataDriverExecutedRedirectURL = "editBroker.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords & "&BrokerID=" & BrokerID & "&DuplicateFound=" & DuplicateFound & "&OldUserName=" & Server.URLEncode(UserName) & "&AttemptUserName=" & Server.URLEncode(AttemptUserName)
		End If
		Response.Redirect DCDataDriverExecutedRedirectURL

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
			.Parameters.Append .CreateParameter("@SyatemLoginID", 200, 1,200,EntitySystemLoginID)
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

		'Updates Broker Information
		DBObjectDestination = "Castles_Brokers"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "BrokerID"
		DataParentNode = ""
		DataExceptions = "<!BrokerID!><!EntityLanguageID!><!EntitySystemLoginID!><!HeaderCount!><!SubHeaderCount!><!NickName!><!UserName!><!Password!><!ShowHelpContent!><!CreateNotes!><!ViewNotes!><!DeleteNotes!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		If DuplicateFound = "N" Then
			DCDataDriverExecutedRedirectURL = "searchBrokers.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		Else
			DCDataDriverExecutedRedirectURL = "editBroker.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords & "&BrokerID=" & BrokerID & "&DuplicateFound=" & DuplicateFound & "&OldUserName=" & Server.URLEncode(UserName) & "&AttemptUserName=" & Server.URLEncode(AttemptUserName)
		End If
		Cnekt = Connect
		
		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Tracks Modification of Broker Account
		EntityID = 7
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,BrokerID,EntityModificationTypeID,IPAddress 

		'Updates System Login Username & Password
		Deleted = "N"
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Connect
		SQLStmt = "UPDATE Castles_SystemLogins SET "
		SQLStmt = SQLStmt & "NickName ='" & NickName & "',"
		SQLStmt = SQLStmt & "UserName ='" & UserName & "',"
		SQLStmt = SQLStmt & "Password ='" & Password & "',"
		SQLStmt = SQLStmt & "LanguageID ='" & EntityLanguageID & "',"
		SQLStmt = SQLStmt & "ShowHelpContent ='" & ShowHelpContent & "',"
		SQLStmt = SQLStmt & "CreateNotes ='" & CreateNotes & "',"
		SQLStmt = SQLStmt & "ViewNotes ='" & ViewNotes & "',"
		SQLStmt = SQLStmt & "Active ='" & Active & "',"
		SQLStmt = SQLStmt & "Deleted ='" & Deleted & "',"
		SQLStmt = SQLStmt & "DeleteNotes ='" & DeleteNotes & "' "
		SQLStmt = SQLStmt & "WHERE SystemLoginID ='" & EntitySystemLoginID & "'"
		Set SQLAction = Conn.Execute(SQLStmt)
		Set SQLAction = Nothing
		Conn.Close

		'Tracks Modification of System Login Info
		EntityID = 6
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,EntitySystemLoginID,EntityModificationTypeID,IPAddress 

		'Updates System Login Access Rights
		Accessable = "Y"
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Connect
		SQLStmt = "DELETE FROM Castles_SystemLoginAccessRights WHERE SystemLoginID ='" & EntitySystemLoginID & "'"
		Set SQLAction = Conn.Execute(SQLStmt)
		Set SQLAction = Nothing
		Conn.Close

		'Tracks Deletion of Broker Access Rights
		EntityID = 9
		EntityModificationTypeID = 3
		EntityModificationLogger SystemLoginID,EntityID,EntitySystemLoginID,EntityModificationTypeID,IPAddress 

		'Header Access Rights Insert
		Count = 1
		For i = 1 to HeaderCount
			If Len(Request.Form("TopNavigationHeaderID" & Count)) <> 0 Then
				TopNavigationHeaderID = Request.Form("TopNavigationHeaderID" & Count)
				TopNavigationSubHeaderID = 0
				Set Conn = Server.CreateObject("ADODB.Connection")
				Conn.Open Connect
				SQLStmt = "INSERT INTO Castles_SystemLoginAccessRights (SystemLoginID,TopNavigationHeaderID,TopNavigationSubHeaderID,LanguageID,Accessable) "
				SQLStmt = SQLStmt & "VALUES (" & "'" & EntitySystemLoginID & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationHeaderID & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationSubHeaderID & "'" 
				SQLStmt = SQLStmt & "," & "'" & EntityLanguageID & "'" 
				SQLStmt = SQLStmt & "," & "'" & Accessable & "'" & ")"
				Set SQLAction = Conn.Execute(SQLStmt)
				Set SQLAction = Nothing
				Conn.Close

				'Tracks Insertion of Broker Header Access Rights
				EntityID = 9
				EntityModificationTypeID = 1
				EntityModificationLogger SystemLoginID,EntityID,EntitySystemLoginID,EntityModificationTypeID,IPAddress 
			End If
			Count = Count + 1
		Next

		'SubHeader Access Rights Insert
		Count = 1
		For i = 1 to SubHeaderCount
			If Len(Request.Form("TopNavigationSubHeaderID" & Count)) <> 0 Then
				TopNavigationHeaderIDForSubHeader = Request.Form("TopNavigationHeaderIDForSubHeader" & Count)
				TopNavigationSubHeaderID = Request.Form("TopNavigationSubHeaderID" & Count)
				Set Conn = Server.CreateObject("ADODB.Connection")
				Conn.Open Connect
				SQLStmt = "INSERT INTO Castles_SystemLoginAccessRights (SystemLoginID,TopNavigationHeaderID,TopNavigationSubHeaderID,LanguageID,Accessable) "
				SQLStmt = SQLStmt & "VALUES (" & "'" & EntitySystemLoginID & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationHeaderIDForSubHeader & "'" 
				SQLStmt = SQLStmt & "," & "'" & TopNavigationSubHeaderID & "'" 
				SQLStmt = SQLStmt & "," & "'" & EntityLanguageID & "'" 
				SQLStmt = SQLStmt & "," & "'" & Accessable & "'" & ")"
				Set SQLAction = Conn.Execute(SQLStmt)
				Set SQLAction = Nothing
				Conn.Close

				'Tracks Insertion of Broker SubHeader Access Rights
				EntityID = 9
				EntityModificationTypeID = 1
				EntityModificationLogger SystemLoginID,EntityID,EntitySystemLoginID,EntityModificationTypeID,IPAddress 
			End If
			Count = Count + 1
		Next

		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'Deletes Broker Information		
		DBObjectDestination = "Castles_Brokers"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "BrokerID"
		DataParentNode = ""
		DataExceptions = ""
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCDataDriverExecutedRedirectURL = "searchBrokers.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		Cnekt = Connect
		EntityID = 7

		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Deletes Broker SystemLogin Information		
		NumberOfRecordsToDelete = Request.Form("NumberOfRecordsToDelete")
		Count = 1
		For i = 1 To NumberOfRecordsToDelete
			If Request.Form("DeleteRecordID" & count) <> "" Then
				UniqueKeyValue = Request.Form("DeleteRecordID" & count)
				Active = "N"
				Deleted = "Y"
				Set Conn = Server.CreateObject("ADODB.Connection")
				Conn.Open Cnekt
				SQLStmt = "UPDATE Castles_SystemLogins SET Active ='" & Active & "',"
				SQLStmt = SQLStmt & "Deleted ='" & Deleted & "' "
				SQLStmt = SQLStmt & "WHERE EntityID = 1 AND EntityPrimaryKeyValue = " & UniqueKeyValue
				Set SQLAction = Conn.Execute(SQLStmt)
				Set SQLAction = Nothing
				Conn.Close

				'Tracks Deletion of Records
				EntityModificationTypeID = 3
				EntityID = 6
				EntityModificationLogger SystemLoginID,EntityID,BrokerID,EntityModificationTypeID,IPAddress 
			End If	
			Count = Count + 1
		Next

		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		Response.Redirect DCDataDriverExecutedRedirectURL
End Select
%>
