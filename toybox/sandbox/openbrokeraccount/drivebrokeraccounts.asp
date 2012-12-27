<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Templates/castlesdcdatadriver.asp" --> 
<%
'Necessary Requests
IPAddress = Request.ServerVariables("REMOTE_ADDR")
Active = "N"
NickName = CleanForDrive(Request.Form("NickName"))
UserName = CleanForDrive(Request.Form("UserName"))
Password = CleanForDrive(Request.Form("Password"))

CreateNotes = CleanForDrive(Request.Form("CreateNotes"))
If Len(CreateNotes) = 0 Then
	CreateNotes = "N"
End If
DeleteNotes = CleanForDrive(Request.Form("DeleteNotes"))
If Len(DeleteNotes) = 0 Then
	DeleteNotes = "N"
End If


DataExceptionsString = ""


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
			
			response.redirect "default.asp?dupusername=Y&CompanyName=" & request("CompanyName") & "&FirstName=" & request("FirstName") & "&MiddleInitial=" & request("MiddleInitial") & "&LastName=" & request("LastName") & "&AddressLine1=" & request("AddressLine1") & "&AddressLine2=" & request("AddressLine2") & "&City=" & request("City") & "&StateProvinceID=" & request("StateProvinceID") & "&ZipPostalCode=" & request("ZipPostalCode") & "&CountryID=" & request("CountryID") & "&TelNumber=" & request("TelNumber") & "&FaxNumber=" & request("FaxNumber") & "&EmailAddress=" & request("EmailAddress") & "&UserName=" & request("UserName")
			'UserName = CleanForDrive(Request.Form("EmailAddress"))
			'AttemptUserName = CleanForDrive(Request.Form("UserName"))
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
		DataExtraFields = "Active<!DCDELIMETER!>" & Active
		Cnekt = Connect
		
		BrokerID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)
		
		'Tracks Modification of Broker Account
		EntityID = 7
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,BrokerID,EntityModificationTypeID,IPAddress 

		'Creates System Login Username & Password
		SystemLoginTypeID = 4
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
			.ActiveConnection = Connect
			.CommandText = "Castles_System_DCdataDriver_SQLInsert_GenerateID"
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


		'If DuplicateFound = "N" Then
			DCDataDriverExecutedRedirectURL = "openbrokeraccountsuccess.asp"
		'Else
			'DCDataDriverExecutedRedirectURL = "default.asp?BrokerID=" & BrokerID & "&DuplicateFound=" & DuplicateFound & "&OldUserName=" & Server.URLEncode(UserName) & "&AttemptUserName=" & Server.URLEncode(AttemptUserName)
		'End If
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

		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'....
End Select
%>
