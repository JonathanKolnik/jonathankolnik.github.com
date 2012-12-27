<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Templates/castlesdcdatadriver.asp" --> 
<%

'Necessary Requests
SystemLoginID = Request.Cookies("SystemLoginID")
IPAddress = Request.ServerVariables("REMOTE_ADDR")
WebSiteContentUniqueID = Request.Form("UniqueID")

'Which DCDataDriverType to use
DCDataDriverType = Request.QueryString("DCDataDriverType")

'Handle different DCDataDriverTypes
Select Case DCDataDriverType
	Case "SQLInsert"
		'...
	Case "SQLUpdate"
		'Updates WebSite Content Information
		DBObjectDestination = "Castles_WebSiteContent"
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
		DCDataDriverExecutedRedirectURL = "searchwebsitecontent.asp"
		Cnekt = Connect
		
		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Tracks Modification of WebSite Content
		EntityID = 20
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,WebSiteContentUniqueID,EntityModificationTypeID,IPAddress 

		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'...		
End Select
%>
