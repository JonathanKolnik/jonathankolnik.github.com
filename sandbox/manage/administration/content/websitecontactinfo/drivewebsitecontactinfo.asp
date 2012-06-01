<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Templates/CastlesDCdataDriver.asp" --> 
<%

'Which DCdataDriverType to use
DCdataDriverType = Request.QueryString("DCdataDriverType")

'Handle different DCdataDriverTypes
Select Case DCdataDriverType
	Case "SQLInsert"
		

	Case "SQLUpdate"

		'Updates WebSiteContent Information
		DBObjectDestination = "Castles_WebSiteContactInfo"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "WebSiteContactInfoID"
		DataParentNode = ""
		DataExceptions = "<!WebSiteContactInfoID!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCdataDriverExecutedRedirectURL = "editwebsitecontactinfo.asp?Update=Y"
		Cnekt = Connect
		
		DCdataDriver DCdataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Tracks Modification of WebSiteContent
		EntityID = 23
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,WebSiteContactInfoID,EntityModificationTypeID,IPAddress 

		Response.Redirect DCdataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		
End Select
%>
