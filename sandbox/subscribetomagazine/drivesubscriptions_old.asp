<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/templates/castlesclientcnekt.asp" -->
<!--#include virtual="/Templates/castlesdcdatadriver.asp" --> 
<!--#include virtual="/Templates/castlesccencrypt.asp" --> 

<%

'Necessary Requests
IPAddress = Request.ServerVariables("REMOTE_ADDR")
CreditCardNumber = Request.Form("CreditCardNumber")
CreditCardTypeID = Request.Form("CreditCardTypeID")
'Response.Write "Initial Credit Card Number = " & CreditCardNumber & "<br><br>"

EncryptedCreditCardNumber = EncryptCreditCardNumber(CreditCardNumber)
DecryptedCreditCardNumber = DecryptCreditCardNumber(EncryptedCreditCardNumber)
'Response.Write "Encrypted Credit Card Number = " & EncryptedCreditCardNumber & "<br><br>"
'Response.Write "Decrypted Credit Card Number = " & DecryptedCreditCardNumber & "<br>"

SubscriptionTransactionDateTime = now
SubscriptionTypeID = Request.Form("SubscriptionTypeID")

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
		DataExceptions = "<!SubscriptionID!><!Same!><!CreditCardNumber!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = "CreditCardNumber<!DCDELIMETER!>" & EncryptedCreditCardNumber & ",SubscriptionTransactionDateTime<!DCDELIMETER!>" & SubscriptionTransactionDateTime & ",FirstIssueMonthYear<!DCDELIMETER!>" & FirstIssueMonthYear & ",ExpirationMonthYear<!DCDELIMETER!>" & ExpirationMonthYear & ",SubscriptionTransactionIP<!DCDELIMETER!>" & IPAddress
		DCDataDriverExecutedRedirectURL = "subscriptionsuccess.asp"
		Cnekt = Connect
		
		SubscriptionID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)
		
		'Tracks Modification of WebSite Content
		EntityID = 7
		EntityModificationTypeID = 1
		EntityModificationLogger SystemLoginID,EntityID,SubscriptionID,EntityModificationTypeID,IPAddress 

		Message = "A subscription has been proccessed through the Castles WebSite..."
		Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
		'objCDOMail.From = "subscriptions@castles.com"
		objCDOMail.From = "info@castlesmag.com"
		objCDOMail.Cc = ""
		objCDOMail.Bcc = "toffling@dreamingcode.com"
		objCDOMail.Subject = "Castles Subscription Submission..."
		objCDOMail.To = "jim_lowenstern@hotmail.com"
		objCDOMail.Body =  Message
		objCDOMail.Send
		Set objCDOMail = Nothing 

		Message = "Thank you for subscribing to the Castles Magazine..."
		Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
		'objCDOMail.From = "subscriptions@castles.com"
		objCDOMail.From = "info@castlesmag.com"
		objCDOMail.Cc = ""
		objCDOMail.Bcc = "toffling@dreamingcode.com"
		objCDOMail.Subject = "Your Castles Magazine Subscription..."
		objCDOMail.To = Request.Form("EmailAddress")
		objCDOMail.Body =  Message
		objCDOMail.Send
		Set objCDOMail = Nothing 

		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLUpdate"
		'...
	Case "SQLMultiDelete"
		'...		
End Select
%>
