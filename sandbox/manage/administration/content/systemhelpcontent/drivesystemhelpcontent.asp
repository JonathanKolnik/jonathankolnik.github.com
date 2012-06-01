<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Templates/castlesdcdatadriver.asp" --> 
<%

'Necessary Requests
SystemLoginID = Request.Cookies("SystemLoginID")
IPAddress = Request.ServerVariables("REMOTE_ADDR")
FilterByTopNavigationHeaderID = Request.QueryString("FilterByTopNavigationHeaderID")
SystemHelpContentUniqueID = Request.Form("UniqueID")

'Which DCDataDriverType to use
DCDataDriverType = Request.QueryString("DCDataDriverType")

'Handle different DCDataDriverTypes
Select Case DCDataDriverType
	Case "SQLInsert"
		'...
	Case "SQLUpdate"
		'Updates System Help Content Information
		DBObjectDestination = "Castles_SystemHelpContent"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "UniqueID"
		DataParentNode = ""
		DataExceptions = "<!UniqueID!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCDataDriverExecutedRedirectURL = "searchsystemhelpcontent.asp?FilterByTopNavigationHeaderID=" & Server.URLEncode(FilterByTopNavigationHeaderID) 
		Cnekt = Connect
		
		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Tracks Modification of Administrator Account
		EntityID = 19
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,SystemHelpContentUniqueID,EntityModificationTypeID,IPAddress 

		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'...		
End Select
%>
