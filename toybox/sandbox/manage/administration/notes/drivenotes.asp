<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Templates/CastlesDCDataDriver.asp" --> 
<%
'Persist Search Values
Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
SearchColumn = Request.QueryString("SearchColumn")
SearchValue = Request.QueryString("SearchValue")

'Necessary Requests
SystemLoginID = Request.Cookies("SystemLoginID")
IPAddress = Request.ServerVariables("REMOTE_ADDR")
EntityID = Request.QueryString("EntityID")
EntityPrimaryKeyValue = Request.QueryString("EntityPrimaryKeyValue")
EntityName = CleanForDrive(Request.QueryString("EntityName"))
PersonalName = CleanForDrive(Request.QueryString("PersonalName"))
PageTopNavigationSubHeaderID = Request.QueryString("PageTopNavigationSubHeaderID")
DataExceptionsString = ""

'Which DCDataDriverType to use
DCDataDriverType = Request.QueryString("DCDataDriverType")

'Handle different DCDataDriverTypes
Select Case DCDataDriverType
	Case "SQLInsert"
		'Inserts Administrator Information
		DBObjectDestination = "Castles_Notes"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "NoteID"
		DataParentNode = ""
		DataExceptions = "<!NoteID!><!EntityName!><!PersonalName!><!Submit!>" & DataExceptionsString 
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = "CreatedDateTime<!DCDELIMETER!>" & now & ",CreatedBySystemLoginID<!DCDELIMETER!>" & SystemLoginID & ",Active<!DCDELIMETER!>Y" & ",EntityID<!DCDELIMETER!>" & EntityID & ",EntityPrimaryKeyValue<!DCDELIMETER!>" & EntityPrimaryKeyValue 
		DCDataDriverExecutedRedirectURL = "noteslist.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue)& "&Page=" & Page & "&TotalRecords=" & TotalRecords & "&EntityID=" & EntityID & "&EntityPrimaryKeyValue=" & EntityPrimaryKeyValue & "&EntityName=" & Server.URLEncode(EntityName) & "&PersonalName=" & Server.URLEncode(PersonalName) & "&PageTopNavigationSubHeaderID=" & PageTopNavigationSubHeaderID
		Cnekt = Connect
		
		NoteID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)

		'Tracks Creation of Note
		EntityID = 17
		EntityModificationTypeID = 1
		EntityModificationLogger SystemLoginID,EntityID,NoteID,EntityModificationTypeID,IPAddress 

		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'Deletes Administrator Information		
		DBObjectDestination = "Castles_Notes"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "NoteID"
		DataParentNode = ""
		DataExceptions = ""
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCDataDriverExecutedRedirectURL = "noteslist.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue)& "&Page=" & Page & "&TotalRecords=" & TotalRecords & "&EntityID=" & EntityID & "&EntityPrimaryKeyValue=" & EntityPrimaryKeyValue & "&EntityName=" & Server.URLEncode(EntityName) & "&PersonalName=" & Server.URLEncode(PersonalName) & "&PageTopNavigationSubHeaderID=" & PageTopNavigationSubHeaderID
		Cnekt = Connect
		EntityID = 17

		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		Response.Redirect DCDataDriverExecutedRedirectURL
End Select
%>
