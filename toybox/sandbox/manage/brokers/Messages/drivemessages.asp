<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/templates/castlesdcdatadriver.asp" --> 
<%
'Persist Search Values
Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
SearchColumn = Request.QueryString("SearchColumn")
SearchValue = Request.QueryString("SearchValue")

'Necessary Requests


DataExceptionsString = ""

'Which DCDataDriverType to use
DCDataDriverType = Request.QueryString("DCDataDriverType")

'Handle different DCDataDriverTypes
Select Case DCDataDriverType
	Case "SQLMultiDelete"
		'Deletes Broker Information		
		DBObjectDestination = "Castles_BrokerMessages"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "BrokerMessageID"
		DataParentNode = ""
		DataExceptions = ""
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCDataDriverExecutedRedirectURL = "searchMessages.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		Cnekt = Connect
		EntityID = 13

		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		Response.Redirect DCDataDriverExecutedRedirectURL
End Select
%>
