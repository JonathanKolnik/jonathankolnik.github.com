<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/templates/castlessystemcnekt.asp" -->
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
SubscriptionID = Request.Form("SubscriptionID")
SubscriptionTypeID = Request.Form("SubscriptionTypeID")

DataExceptionsString = ""
SubscriptionTransactionDateTime = now
ClientSideOrSystem = "S"

'Which DCDataDriverType to use
DCDataDriverType = Request.QueryString("DCDataDriverType")

'Handle different DCDataDriverTypes
Select Case DCDataDriverType
	Case "SQLInsert"
		'calculate the first issue and subscription expiration
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_SubscriptionType_PeriodMonths"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@SubscriptionTypeID", 200, 1,200,SubscriptionTypeID)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set SubscriptionPeriod = .Execute()
		End With
		Set Command1 = Nothing
		
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_FirstIssueCalculator"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set FirstIssueCalc = .Execute()
		End With
		Set Command1 = Nothing
		
		SubscriptionPeriodMonths = SubscriptionPeriod.Fields.Item("SubscriptionPeriodMonths").Value
		DateSpan = FirstIssueCalc.Fields.Item("DateSpan").Value
		CutoffDay = FirstIssueCalc.Fields.Item("CutoffDay").Value
		
		CurrentMonth = DatePart("m",SubscriptionTransactionDateTime)
		AfterSpanDate = DateAdd("d",DateSpan,SubscriptionTransactionDateTime)
		AfterSpanYear = DatePart("yyyy",AfterSpanDate)
		AfterSpanMonth = DatePart("m",AfterSpanDate)
		AfterSpanDay = DatePart("d",AfterSpanDate)
		if AfterSpanDay < CutoffDay then
			FirstIssueMonthYear = AfterSpanMonth & "/1/" & AfterSpanYear
		elseif AfterSpanDay >= CutoffDay then
			FirstIssueMonthYear = (AfterSpanMonth+1) & "/1/" & AfterSpanYear
		end if
		ExpirationMonthYear = DateAdd("m",SubscriptionPeriodMonths,FirstIssueMonthYear)
		
		
		'Inserts Subscription Information
		DBObjectDestination = "Castles_Subscriptions"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "SubscriptionID"
		DataParentNode = ""
		DataExceptions = "<!UseBilling!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = "SubscriptionTransactionDateTime<!DCDELIMETER!>" & SubscriptionTransactionDateTime & ",ClientSideOrSystem<!DCDELIMETER!>" & ClientSideOrSystem & ",FirstIssueMonthYear<!DCDELIMETER!>" & FirstIssueMonthYear & ",ExpirationMonthYear<!DCDELIMETER!>" & ExpirationMonthYear & ",SubscriptionTransactionIP<!DCDELIMETER!>" & IPAddress
		Cnekt = Connect
		
		SubscriptionID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)
		
		'Tracks Modification of Subscription Account
		EntityID = 11
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,SubscriptionID,EntityModificationTypeID,IPAddress 

		DCDataDriverExecutedRedirectURL = "searchSubscriptions.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		
		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLUpdate"
		'Updates Subscription Information
		DBObjectDestination = "Castles_Subscriptions"
		FormType = "Request.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "SubscriptionID"
		DataParentNode = ""
		DataExceptions = "<!SubscriptionID!><!UseBilling!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCDataDriverExecutedRedirectURL = "searchSubscriptions.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		
		Cnekt = Connect
		
		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Tracks Modification of Subscription Account
		EntityID = 11
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,SubscriptionID,EntityModificationTypeID,IPAddress 

		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'Deletes Subscription Information		
		DBObjectDestination = "Castles_Subscriptions"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "SubscriptionID"
		DataParentNode = ""
		DataExceptions = ""
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCDataDriverExecutedRedirectURL = "searchSubscriptions.asp?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		Cnekt = Connect
		EntityID = 11

		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		Response.Redirect DCDataDriverExecutedRedirectURL
End Select
%>
