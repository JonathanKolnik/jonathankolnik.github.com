<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../templates/castlesclientcnekt.asp" -->
<!--#include file="../Templates/castlesdcdatadriver.asp" --> 
<!--#include file="../Templates/castlesccencrypt.asp" --> 
<%
'Displays WebSite Content
WebSiteContentID = 6
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_ClientSide_WebSiteContent_Detail"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4) 
	.Parameters.Append .CreateParameter("@WebSiteContentID", 200, 1,200,WebSiteContentID)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,LanguageID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set WebSiteContent = .Execute()
End With
Set Command1 = Nothing

If Not WebSiteContent.EOF Then
	WebSiteContentName = WebSiteContent.Fields.Item("WebSiteContentName").Value
	WebSiteContentCaption1 = WebSiteContent.Fields.Item("WebSiteContentCaption1").Value
	WebSiteContentCaptionHeader1 = WebSiteContent.Fields.Item("WebSiteContentCaptionHeader1").Value

	If Len(WebSiteContentCaption1) <> 0 Then
		WebSiteContentCaption1 = Replace(WebSiteContentCaption1,vbcrlf,"<br>")
	End If
	
	WebSiteContentBody1 = WebSiteContent.Fields.Item("WebSiteContentBody1").Value
	If Len(WebSiteContentBody1) <> 0 Then
		WebSiteContentBody1 = Replace(WebSiteContentBody1,vbcrlf,"<br>")
	End If
End If
%>
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
		DataExceptions = "<!SubscriptionID!><!Same!><!hidBillingCountryID!><!hidShippingCountryID!><!hidCreditCardTypeID!><!hidShippingStateProvinceID!><!hidBillingStateProvinceID!><!hidSubscriptionChargeSelected!><!hidSubscriptionCharge[]!><!CreditCardNumber!><!Submit!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = "CreditCardNumber<!DCDELIMETER!>" & EncryptedCreditCardNumber & ",SubscriptionTransactionDateTime<!DCDELIMETER!>" & SubscriptionTransactionDateTime & ",FirstIssueMonthYear<!DCDELIMETER!>" & FirstIssueMonthYear & ",ExpirationMonthYear<!DCDELIMETER!>" & ExpirationMonthYear & ",SubscriptionTransactionIP<!DCDELIMETER!>" & IPAddress
		DCDataDriverExecutedRedirectURL = "subscriptionsuccess.asp"
		Cnekt = Connect
		
		SubscriptionID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)
		'**********************************************************************************************
		''SubscriptionID=123
		'**********************************************************************************************
		'Tracks Modification of WebSite Content
		EntityID = 7
		EntityModificationTypeID = 1
		EntityModificationLogger SystemLoginID,EntityID,SubscriptionID,EntityModificationTypeID,IPAddress 

		''Message = "A subscription has been proccessed through the Castles WebSite..."
		''Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
		'objCDOMail.From = "subscriptions@castles.com"
		''objCDOMail.From = "info@castlesmag.com"
		''objCDOMail.Cc = ""
		''objCDOMail.Bcc = "toffling@dreamingcode.com"
		''objCDOMail.Subject = "Castles Subscription Submission..."
		''objCDOMail.To = "jim_lowenstern@hotmail.com"
		''objCDOMail.Body =  Message
		''objCDOMail.Send
		''Set objCDOMail = Nothing 
	'**********************************************************************************************************
		


    dim reqFieldsOK
	dim retVal
	dim noErrors
	dim errorMessage

    'the COM client
    dim eClient

	dim accountToken

	dim orderID
	dim chargeTotal
	dim chargeType
	dim creditCardNumber
	dim cvv
	dim tcc
	dim expireMonth
	dim expireYear
	dim orderDescription
	dim orderUserID
	dim shippingCharge
	dim taxAmount
	dim doTransactionOnAuthenticationInconclusive

	dim billFirstName
	dim billMiddleName
	dim billLastName
	dim billAddressOne
	dim billAddressTwo
	dim billCity
	dim billStateOrProvince
	dim billCountryCode
	dim billZipOrPostalCode
	dim billEmail
	dim billCompany
	dim billPhone
	dim billFax
	dim billNote

	dim shipFirstName
	dim shipMiddleName
	dim shipLastName
	dim shipAddressOne
	dim shipAddressTwo
	dim shipCity
	dim shipStateOrProvince
	dim shipCountryCode
	dim shipZipOrPostalCode
	dim shipPhone
	dim shipFax
	dim shipEmail
	dim shipCompany
	dim shipNote

	dim paramPARes
	dim paramPAResEncoded
	dim paramMD
	dim passedPayerAuthentication
	dim isEnrolled

	' Create variables for each of the response fields to
	' print to the receipt page.
	dim respResponseCode
	dim respResponseCodeText
	dim respTimeStamp
	dim respTimeStampString
	dim respOrderID
	dim respReferenceID
	dim respISOCode
	dim respBankApprovalCode
	dim respBankTransactionID
	dim respBatchID
	dim respCreditCardVerificationValueResponse
	dim respAVSCode

	' If payer authentication was attempted, the CreditCardResponse
	' will contain a PayerAuthenticationResponse object with
	' details of the authentication results
	dim authRespResponseCode
	dim authRespResponseCodeString
	dim authRespResponseCodeText
	dim authRespCAVV
	dim authRespXID
	dim authRespTCC
	dim authRespTCCString
	dim authRespTimeStamp


	'instantiate the COM client
	set eClient = Server.CreateObject("Paygateway.EClient.1")

	'initialize ALL the variables
	reqFieldsOK = true
	errorMessage = ""
	passedPayerAuthentication = false

	respResponseCode = 0
	respResponseCodeText = ""
	respTimeStamp = ""
	respOrderID = ""
	respReferenceID = ""
	respISOCode = ""
	respBankApprovalCode = ""
	respBankTransactionID = ""
	respBatchID = ""
	respCreditCardVerificationValueResponse = ""
	respAVSCode = ""

	authRespResponseCode = 0
	authRespResponseCodeString = ""
	authRespResponseCodeText = ""
	authRespCAVV = ""
	authRespXID = ""
	authRespTCC = 0
	authRespTCCString = ""
	authRespTimeStamp = ""
	
	accountToken = "3CF0CE4ECE60580B338EEEF08755AA63CBC8289E42B0C1F6E5A94E2585D85AEA9D50B8BDFF08F01D"
	'accountToken = Application("ACCOUNT_TOKEN")

	paramPARes			= Request("PaRes")
	paramPAResEncoded	= Request("PaResEncoded")
	paramMD				= Request("MD")

'Store posted variables
	
	chargeTotal = Request.Form("hidSubscriptionChargeSelected")
	chargeType = "AUTH"'"SALE"'Request.Form("hidCreditCardTypeID")
	creditCardNumber = Request.Form("CreditCardNumber")'"4242424242424242"'
	cvv = Request.Form("cvv")
	tcc = Request.Form("CreditCardCSCCode") '("transaction_condition_code")
	expireMonth = Request.Form("CreditCardExpirationMonth").item
	expireYear = Request.Form("CreditCardExpirationYear").item
	'orderDescription = Request.Form("order_description")
	orderID = SubscriptionID 'Request.Form("order_id")
	'orderUserID = Request.Form("order_user_id")
	shippingCharge = Request.Form("shipping_charge")
	taxAmount = Request.Form("tax_amount")

	billAddressOne = Request.Form("BillingAddressLine1")
	billAddressTwo = Request.Form("BillingAddressLine2")
	billCity = Request.Form("BillingCity")
	'billCompany = Request.Form("bill_company")
	'????????
	billCountryCode = Request.Form("hidBillingCountryID")
	'????????BillingCountryID
	'????????SubscriberComments-->?
	'????????CreditCardHolderName -----> Full Name
	'billCustomerTitle = Request.Form("bill_customer_title")
	billEmail = Request.Form("EmailAddress")
	
	billFirstName = Request.Form("BillingFirstName")
	billLastName = Request.Form("BillingLastName")
	'billMiddleName = Request.Form("bill_middle_name")
	billNote = Request.Form("SubscriberComments")
	billPhone = Request.Form("TelNumber")
	billStateOrProvince = Request.Form("hidBillingStateProvinceID")
	billZipOrPostalCode = Request.Form("BillingZipPostalCode")

	shipAddressOne = Request.Form("ShippingAddressLine1")
	shipAddressTwo = Request.Form("ShippingAddressLine2")
	shipCity = Request.Form("ShippingCity")
	'shipCompany = Request.Form("ship_company")
	shipCountryCode = Request.Form("hidShippingCountryID")
	'shipCustomerTitle = Request.Form("ship_customer_title")
	shipEmail = Request.Form("EmailAddress")
	shipFirstName = Request.Form("ShippingFirstName")
	shipLastName = Request.Form("ShippingLastName")
	'shipMiddleName = Request.Form("ship_middle_name")
	shipNote = Request.Form("SubscriberComments")
	shipPhone = Request.Form("TelNumber")
	shipStateOrProvince = Request.Form("hidShippingStateProvinceID")
	shipZipOrPostalCode = Request.Form("ShippingZipPostalCode")
'response.Write(shipEmail & "shipEmail:")
'response.End
	
	'response.Write eClient.GetCreditCardNumber

	' Read in variables from the session
''	isEnrolled = Session("use_payer_authentication")
''	authenticationTransactionID = Session("authentication_transaction_id")
''	chargeTotal = Session("charge_total")
'	chargeType = Session("charge_type")
'	creditCardNumber = Session("credit_card_number")
'	cvv = Session("cvv")
'	tcc = Session("transaction_condition_code")
'	expireMonth = Session("expire_month")
'	expireYear = Session("expire_year")
'	orderDescription = Session("order_description")
'	orderID = Session("order_id")
'	orderUserID = Session("order_user_id")
'	shippingCharge = Session("shipping_charge")
'	taxAmount = Session("tax_amount")
'
''	billAddressOne = Session("bill_address_one")
'	billAddressTwo = Session("bill_address_two")
'	billCity = Session("bill_city")
'	billCompany = Session("bill_company")
'	billCountryCode = Session("bill_country_code")
'	billCustomerTitle = Session("bill_customer_title")
'	billEmail = Session("bill_email")
'	billFirstName = Session("bill_first_name")
'	billLastName = Session("bill_last_name")
'	billMiddleName = Session("bill_middle_name")
'	billNote = Session("bill_note")
'	billPhone = Session("bill_phone")
'	billStateOrProvince = Session("bill_state_or_province")
'	billZipOrPostalCode = Session("bill_zip_or_postal_code")
'
'	shipAddressOne = Session("ship_address_one")
'	shipAddressTwo = Session("ship_address_two")
'	shipCity = Session("ship_city")
'	shipCompany = Session("ship_company")
'	shipCountryCode = Session("ship_country_code")
'	shipCustomerTitle = Session("ship_customer_title")
'	shipEmail = Session("ship_email")
'	shipFirstName = Session("ship_first_name")
'	shipLastName = Session("ship_last_name")
'	shipMiddleName = Session("ship_middle_name")
''	shipNote = Session("ship_note")
''	shipPhone = Session("ship_phone")
''	shipStateOrProvince = Session("ship_state_or_province")
''	shipZipOrPostalCode = Session("ship_zip_or_postal_code")

''	doTransactionOnAuthenticationInconclusive = Session("do_transaction_on_authentication_inconclusive")



	'set methods

	if reqFieldsOK then
		retVal = eClient.SetCreditCardNumber(creditCardNumber)
		'required field
		if (retVal = 0) then
			errorMessage = errorMessage & "invalid credit card number <br>"
			reqFieldsOK = false
		end if
	end if

	if reqFieldsOK then
		retVal = eClient.SetExpireMonth(expireMonth)
		'required field
		if (retVal = 0) then
			errorMessage = errorMessage & "invalid expire month <br>"
			reqFieldsOK = false
		end if
	end if

	if reqFieldsOK then
		retVal = eClient.SetExpireYear(expireYear)
		'required field
		if (retVal = 0) then
			errorMessage = errorMessage & "invalid expire year <br>"
			reqFieldsOK = false
		end if
	end if

	if reqFieldsOK then
		'ChargeType is one of SALE, AUTH, CAPTURE, VOID, CREDIT
		retVal = eClient.SetChargeType(chargetype)
		'required field
		if (retVal = 0) then
			errorMessage = errorMessage & "invalid charge type <br>"
			reqFieldsOK = false
		end if
	end if

	if reqFieldsOK then
		retVal = eClient.SetChargeTotal(chargeTotal)
		'required field
		if (retVal = 0) then
			errorMessage = errorMessage & "invalid charge total <br>"
			reqFieldsOK = false
		end if
	end if

	if reqFieldsOK then
		if( not "" = orderID) then
			retVal = eClient.SetOrderId(orderID)
		end if
		'required field
		if (retVal = 0) then
			errorMessage = errorMessage & "invalid order ID <br>"
			reqFieldsOK = false
		end if
	end if

	if reqFieldsOK then
		'none of these are required
		retVal = eClient.SetCreditCardVerificationNumber(cvv)
		retVal = eClient.SetTaxAmount(taxAmount)
		retVal = eClient.SetShippingCharge(shippingCharge)
		retVal = eClient.SetTransactionConditionCode(tcc)
		retVal = eClient.SetBillFirstName(billFirstName)
		retVal = eClient.SetBillMiddleName(billMiddleName)
		retVal = eClient.SetBillLastName(billLastName)
		retVal = eClient.SetBillAddressOne(billAddressOne)
		retVal = eClient.SetBillAddressTwo(billAddressTwo)
		retVal = eClient.SetBillCity(billCity)
		retVal = eClient.SetBillCompany(billCompany)
		retVal = eClient.SetBillStateOrProvince(billStateOrProvince)
		retVal = eClient.SetBillPostalCode(billZipOrPostalCode)
		retVal = eClient.SetBillCountryCode(billCountryCode)
		retVal = eClient.SetBillPhone(billPhone)
		retVal = eClient.SetBillNote(billNote)
		retVal = eClient.SetShipFirstName(shipFirstName)
		retVal = eClient.SetShipMiddleName(shipMiddleName)
		retVal = eClient.SetShipLastName(shipLastName)
		retVal = eClient.SetShipAddressOne(shipAddressOne)
		retVal = eClient.SetShipAddressTwo(shipAddressTwo)
		retVal = eClient.SetShipCity(shipCity)
		retVal = eClient.SetShipCompany(shipCompany)
		retVal = eClient.SetShipStateOrProvince(shipStateOrProvince)
		retVal = eClient.SetShipPostalCode(shipZipOrPostalCode)
		retVal = eClient.SetShipCountryCode(shipCountryCode)
		retVal = eClient.SetShipPhone(shipPhone)
		retVal = eClient.SetShipNote(shipNote)
		retVal = eClient.SetShipEmail(shipEmail)
		retVal = eClient.SetBillEmail(billEmail)		

	end if

	if( not "" = paramPARes ) then

		eClient.SetAuthenticationPayload( Request("PaRes") )

		if( not "" = authenticationTransactionID ) then
			eClient.SetAuthenticationTransactionId( authenticationTransactionID )
		end if

		eClient.SetDoTransactionOnAuthenticationInconclusive( doTransactionOnAuthenticationInconclusive )

		' You must set the transaction condition code to indicate that you
		' want to perform payer authentication.
		eClient.SetTransactionConditionCode( 4 )

	else

		'Bypassing payer authentication.

	end if

	if( "" = errorMessage ) then

		retVal = eClient.DoTransaction("trans_key", accountToken)

		if (retVal = 0)  then
			eClient.GetErrorString errorMessage
			noErrors = false
		else
			noErrors = true
		end if

		if noErrors then
			eClient.GetResponseCode respResponseCode
			eClient.GetResponseCodeText respResponseCodeText
			eClient.GetTimeStamp respTimeStamp
			eClient.GetOrderId respOrderID
			eClient.GetReferenceId respReferenceID
			eClient.GetIsoCode respISOCode
			eClient.GetBankApprovalCode respBankApprovalCode
			eClient.GetBankTransactionId respBankTransactionID
			eClient.GetBatchId respBatchID
			eClient.GetCreditCardVerificationResponse respCreditCardVerificationValueResponse
			eClient.GetAVSCode respAVSCode
			eClient.GetTransactionConditionCode tcc

			' -- Retreiving Authentication Responses --
			eClient.GetAuthenticationResponseCode authRespResponseCode
			eClient.GetAuthenticationResponseCodeText authRespResponseCodeText
			eClient.GetAuthenticationTimeStamp authRespTimeStamp
			eClient.GetAuthenticationCAVV authRespCAVV
			eClient.GetAuthenticationXID authRespXID
			eClient.GetAuthenticationTransactionConditionCode authRespTCC
			
			session("respResponseCodeText") = respResponseCodeText
			session("respTimeStamp") =	respTimeStamp
			session("respBankApprovalCode") = respBankApprovalCode
			session("respBankTransactionID") =respBankTransactionID 
			session("respReferenceID") = respReferenceID
			session("chargeTotal") = chargeTotal

		
		if instr(respResponseCodeText,"Declined") or instr(respResponseCodeText,"declined") then
			resAcceptedOrDeclined = "D"
		else
			resAcceptedOrDeclined = "A"
		end if
			

			Set Command1 = Server.CreateObject("ADODB.Command")
			With Command1	
				.ActiveConnection = Connect
				.CommandText = "Castles_Update_Subscriptions"
				.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4) 
				.Parameters.Append .CreateParameter("@SubscriptionID", 3, 1,4,respOrderID)
				.Parameters.Append .CreateParameter("@VerisignResultCode", 200, 1,50,respResponseCodeText)
				.Parameters.Append .CreateParameter("@AcceptedOrDeclined", 200, 1,1, resAcceptedOrDeclined)
				.CommandType = 4
				.CommandTimeout = 0
				.Prepared = True
				Set WebSiteContent = .Execute()
			End With
			Set Command1 = Nothing
			
			if( "1" = authRespResponseCode  ) then
				passedPayerAuthentication = true
			end if

%>
<head>

<title>Castles Magazine - The International Magazine for Distinctive Properties</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.CastlesTextBlack {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 400; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBlackBold {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextWhite {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 400; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextWhiteHeader {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextWhiteBold {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextNav {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #333333; text-decoration: none}
.CastlesTextBody {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 400; font-variant: normal; text-transform: none; color: #666633; text-decoration: none}
.CastlesTextBodyBold {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #666633; text-decoration: none}

A.normal:link    { text-decoration: none; color: "#666633"; font-weight: 800}
A.normal:visited { text-decoration: none; color: "#666633"; font-weight: 800}
A.normal:active  { text-decoration: none; color: "#666633"; font-weight: 800}
A.normal:hover   { text-decoration: underline; color: "#666633"; font-weight: 800}

A.white:link    { text-decoration: none; color: "#FFFFFF"; font-weight: 800}
A.white:visited { text-decoration: none; color: "#FFFFFF"; font-weight: 800}
A.white:active  { text-decoration: none; color: "#FFFFFF"; font-weight: 800}
A.white:hover   { text-decoration: underline; color: "#FFFFFF"; font-weight: 800}

A.black:link    { text-decoration: none; color: "#333333"; font-weight: 800}
A.black:visited { text-decoration: none; color: "#333333"; font-weight: 800}
A.black:active  { text-decoration: none; color: "#333333"; font-weight: 800}
A.black:hover   { text-decoration: underline; color: "#333333"; font-weight: 800}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="3" topmargin="3" marginwidth="3" marginheight="3">
<table width="750" border="0" cellspacing="0" cellpadding="0" ID="Table4">
  <tr> 
    <td> 
      <table width="750" border="0" cellspacing="0" cellpadding="0" ID="Table5">
        <tr> 
          <td bgcolor="#CE6F19" height="10" width="350"><img src="../images/clear10pixel.gif" width="10" height="10"></td>
          <td width="400" height="10"><img src="../images/tagline.GIF" width="387" height="10"></td>
        </tr>
        <tr> 
          <td width="350"><a href="/default.asp"><img src="../images/cstles_logo.gif" width="338" height="60" border="0" alt="Castles Magazine"></a></td>
          <td width="400" valign="bottom" align="right"> 
            <table width="400" cellspacing="0" cellpadding="0" border="0" ID="Table6">
              <tr> 
                <td width="300" align="right" class="CastlesTextBodyBold"><a href="javascript:BrokerLogin()" class="black">&gt; 
                  Submit FREE Online Listing/Magazine Ad (Brokers)</a></td>
                <td width="100" align="right" class="CastlesTextBodyBold"><a href="../contactcastles/default.asp" class="black">&gt; 
                  Contact Us</a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="750" border="0" cellspacing="0" cellpadding="0" ID="Table7">
        <tr> 
          <td width="150" valign="top" bgcolor="#CCCCCC"> 
            <table width="150" border="0" cellspacing="0" cellpadding="0" ID="Table9">
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3" ID="Table10">
                    <tr> 
                      <td class="CastlesTextNav"><a href="/default.asp" class="black">Home</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#A0A0A0"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3" ID="Table11">
                    <tr> 
                      <td class="CastlesTextNav"><a href="/Listings/default.asp" class="black">Search 
                        Properties</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="/FeaturedListings/default.asp" class="black">Featured 
                        Properties</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="javascript:BrokerLogin()" class="black">Broker 
                        Log-in / Place Ad</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3" ID="Table12">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../aboutcastles" class="black">About 
                        Castles</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../contactcastles" class="black">Contact 
                        Castles</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3" ID="Table13">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../subscribetomagazine" class="black">Subscribe 
                        to Magazine</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3" ID="Table14">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../openbrokeraccount" class="black">Open 
                        a Broker Account</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3" ID="Table15">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../advertisingrates" class="black">Advertising 
                        Rates</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>

              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3" ID="Table16">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../submissionguidelines" class="black">Submission 
                        Guidelines</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>

              <tr bgcolor="#CCCCCC"> 
                <td width="10">&nbsp;</td>
                <td width="130" height="50">&nbsp;</td>
                <td width="10">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
          <td bgcolor="#CCCCCC" width="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
          <td width="595" valign="top" colspan="4"><!-- #BeginEditable "body" -->
            <table width="595" border="0" cellspacing="0" cellpadding="0" ID="Table17">
              <tr> 
                <td width="395" valign="top"> 
                  <table width="395" border="0" cellspacing="0" cellpadding="0" ID="Table18">
                    <tr bgcolor="#CC6600"> 
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                      <td width="375"> 
                        <table width="355" border="0" cellspacing="0" cellpadding="3" height="20" ID="Table19">
                          <tr> 
                            <td class="CastlesTextWhiteBold"><%=WebSiteContentName%></td>
                          </tr>
                        </table>
                      </td>
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                    </tr>
                    <tr> 
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                      <td width="375" valign="top"> 
                        <table width="375" border="0" cellspacing="0" cellpadding="3" ID="Table20">
                          <tr> 
                            <td class="CastlesTextBody">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td class="CastlesTextBody"><%=WebSiteContentBody1%> </td>
                          </tr>
                          <tr>
                            <td class="CastlesTextBody">
                            <!--**********************************Transaction result*****************************************-->
                            
                            <br>
			<table align="center" border = "0" cellspacing = "0" cellpadding = "0" ID="Table1">
			<tr align="left"> 
				<td colspan="2" class="CastlesTextBody"><b>Successful 
				Subscription </b></td>
			</tr>
			<tr align="left"> 
				<td colspan="2" class="CastlesTextBody">
				Thank you for subscribing to Castles 
				Magazine. A confirmation email has been 
				sent to your specified email address with 
				your subscription information. If you 
				have any questions regarding your subscription 
				please contact us via email at subscriptions@castles.com 
				or via telephone at 1-800-555-5785.<br>
				<br>Thanks,<br><br>
				Castles Unlimited<br><br>
				</td>
			</tr>
			<tr><td>&nbsp;</td></tr>	
			  <tr>
				<td align="left" valign="top">&nbsp;</td>
			    <td align="center" valign="top" class="CastlesTextNav">Transaction Results</td>
				<td align="right" valign="top">&nbsp;</td>
			  </tr>
			</table>
			
			<center>
			<TABLE BORDER="1" ID="Table2"><TR ALIGN="CENTER">
			  <TABLE CELLPADDING="2" CELLSPACING="2" BORDER="0" class="CastlesTextBody">
			<!--	<tr class = "header">
				  <td colspan=2>&nbsp;Transaction Response Fields</td>
				</tr>
			<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Response Code:</TH>
				  <TD width="300">&nbsp;<%=respResponseCode%></TD>
				</TR>
			-->
				<TR ALIGN="LEFT">
				  <TH align="right" width="50%" valign="top" class="CastlesTextBody">Response Text:</TH>
				  <TD width="50%" class="CastlesTextBody">&nbsp;<%=respResponseCodeText%>
					<input name="hidrespResponseCodeText" type="hidden" id="hidrespResponseCodeText">
				  </TD>
				</TR>
			<!--	<TR ALIGN="LEFT">
				  <TH align="right" class="CastlesTextBody" valign="top">Order ID:</TH>
				  <TD class="CastlesTextBody">&nbsp;<%=respOrderID%></TD>
				</TR>
			-->
			<TR ALIGN="LEFT">
				  <TH align="right" class="CastlesTextBody" valign="top">Timestamp:</TH>
				  <TD class="CastlesTextBody">&nbsp;<%=respTimeStamp%><input name="hidrespTimeStamp" type="hidden" id="hidrespTimeStamp">
				  </TD>
				</TR>
			
			<!--<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">AVS Response Code:</TH>
				  <TD width="300">&nbsp;<%=respAVSCode%></TD>
				</TR>
			-->
				<TR ALIGN="LEFT">
				  <TH align="right" class="CastlesTextBody" valign="top">Bank Approval Code:</TH>
				  <TD class="CastlesTextBody">&nbsp;<%=respBankApprovalCode%><input name="hidrespBankApprovalCode" type="hidden" id="hidrespBankApprovalCode">
				  </TD>
				</TR>
			<TR ALIGN="LEFT">
				  <TH align="right" class="CastlesTextBody" valign="top">Bank Transaction ID:</TH>
				  <TD class="CastlesTextBody">&nbsp;<%=respBankTransactionID%><input name="hidrespBankTransactionID" type="hidden" id="hidrespBankTransactionID">
				  </TD>
				</TR>			
			<!--<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Batch ID:</TH>
				  <TD width="300">&nbsp;<%=respBatchID%></TD>
				</TR>
			-->
			<TR ALIGN="LEFT">
				  <TH class="CastlesTextBody" align="right" valign="top">Reference ID:</TH>
				  <TD class="CastlesTextBody">&nbsp;<%=respReferenceID%><input name="hidrespReferenceID" type="hidden" id="hidrespReferenceID">
				  </TD>
				</TR>
			
			<!--<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Credit Card Verification Value Response:</TH>
				  <TD width="300">&nbsp;<%=respCreditCardVerificationValueResponse%></TD>
				</TR>
			-->
			<!--<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">ISO Code:</TH>
				  <TD width="300">&nbsp;<%=respISOCode%></TD>
				</TR>
			-->
			<!--  </TABLE>
			  <p>&nbsp;</p>
			  <TABLE CELLPADDING="2" CELLSPACING="2" BORDER="0" width="500" ID="Table4">
				<tr class = "header">
				  <td colspan=2>&nbsp;Authentication Response Fields</td>
				</tr>
				<tr align="LEFT">
				  <th valign="top" align="right" width="200">Enrolled in Payer Authentication:</th>
				  <TD width="300">&nbsp;<%=isEnrolled%></td>
				</tr>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">Passed Payer Authentication:</th>
				  <TD width="300">&nbsp;<%=passedPayerAuthentication%></td>
				</tr>
<%
			if( isEnrolled ) then
%>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Response Code:</TH>
				  <TD width="300">&nbsp;<%=authRespResponseCode%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Response Code Text:</TH>
				  <TD width="300">&nbsp;<%=authRespResponseCodeText%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Time Stamp:</TH>
				  <TD width="300">&nbsp;<%=authRespTimeStamp%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">CAVV:</TH>
				  <TD width="300">&nbsp;<%=authRespCAVV%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">XID:</TH>
				  <TD width="300">&nbsp;<%=authRespXID%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Transaction Condition Code:</TH>
				  <TD width="300">&nbsp;<%=authRespTCC%></TD>
				</TR>
<%
			end if		' End enrolled section
%>
			  </table>
			  <p>&nbsp;</p>
			  <table cellpadding="2" cellspacing="2" border="0" width="500" ID="Table5">
				<tr class = "header">
				  <td colspan=2>&nbsp;Financial Details</td>
				</tr>
			-->
				<tr align="LEFT">
				  <th valign="top" align="right" class="CastlesTextBody">Charge Total:</th>
				  <TD class="CastlesTextBody">&nbsp;<%=chargeTotal%><input name="hidchargeTotal" type="hidden" id="hidchargeTotal">
				  </td>
				</tr>
			<!--	<tr align="LEFT">
				  <th align="right" width="200" valign="top">Tax Amount:</th>
				  <TD width="300">&nbsp;<%=taxAmount%></td>
				</tr>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">Shipping Charge:</th>
				  <TD width="300">&nbsp;<%=shippingCharge%></td>
				</tr>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">Charge Type:</th>
				  <TD width="300">&nbsp;<%=chargeType%></td>
				</tr>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">Credit Card
					Number:</th>
				  <TD width="300">&nbsp;<%=creditCardNumber%></td>
				</tr>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">Expiry Date (MM/YYYY):</th>
				  <TD width="300">&nbsp;<%=expireMonth%>/<%=expireYear%></td>
				</tr>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">Transaction Condition Code:</th>
				  <TD width="300">&nbsp;<%=tcc%></td>
				</tr>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">Details:</th>
				  <TD width="300">&nbsp;<%=orderDescription%></td>
				</tr>
			  </table>
			  <p>&nbsp;</p>
			  <TABLE CELLPADDING="2" CELLSPACING="2" BORDER="0" width="500" ID="Table6">
				<tr class = "header">
				  <td colspan=2>&nbsp;Billing Information</td>
				</tr>
				<TR ALIGN="LEFT">
				  <TH valign="top" align="right" width="200">Customer Title:</TH>
				  <TD width="300">&nbsp;<%=billCustomerTitle%></TD>

				</TR>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">First Name:</th>
				  <TD width="300">&nbsp;<%=billFirstName%></td>
				</tr>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Last Name:</TH>
				  <TD width="300">&nbsp;<%=billLastName%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Middle Name:</TH>
				  <TD width="300">&nbsp;<%=billMiddleName%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Company:</TH>
				  <TD width="300">&nbsp;<%=billCompany%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Address One:</TH>
				  <TD width="300">&nbsp;<%=billAddressOne%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Address Two:</TH>
				  <TD width="300">&nbsp;<%=billAddressTwo%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">City:</TH>
				  <TD width="300">&nbsp;<%=billCity%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">State or Province:</TH>
				  <TD width="300">&nbsp;<%=billStateOrProvince%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Country Code:</TH>
				  <TD width="300">&nbsp;<%=billCountryCode%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Zip or Postal Code:</TH>
				  <TD width="300">&nbsp;<%=billZipOrPostalCode%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Phone:</TH>
				  <TD width="300">&nbsp;<%=billPhone%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Email:</TH>
				  <TD width="300">&nbsp;<%=billEmail%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Note:</TH>
				  <TD width="300">&nbsp;<%=billNote%></TD>
				</TR>
			  </TABLE>
			<p>&nbsp;</p>
			  <TABLE CELLPADDING="2" CELLSPACING="2" BORDER="0" width="500" ID="Table7">
				<tr class = "header">
				  <td colspan=2>&nbsp;Shipping Information</td>
				</tr>
				<TR ALIGN="LEFT">
				  <TH valign="top" align="right" width="200">Customer Title:</TH>
				  <TD width="300">&nbsp;<%=shipCustomerTitle%></TD>
				</TR>
				<tr align="LEFT">
				  <th align="right" width="200" valign="top">First Name:</th>
				  <TD width="300">&nbsp;<%=shipFirstName%></td>
				</tr>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Last Name:</TH>
				  <TD width="300">&nbsp;<%=shipLastName%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Middle Name:</TH>
				  <TD width="300">&nbsp;<%=shipMiddleName%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Company:</TH>
				  <TD width="300">&nbsp;<%=shipCompany%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Address One:</TH>
				  <TD width="300">&nbsp;<%=shipAddressOne%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Address Two:</TH>
				  <TD width="300">&nbsp;<%=shipAddressTwo%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">City:</TH>
				  <TD width="300">&nbsp;<%=shipCity%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">State or Province:</TH>
				  <TD width="300">&nbsp;<%=shipStateOrProvince%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Country Code:</TH>
				  <TD width="300">&nbsp;<%=shipCountryCode%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Zip or Postal Code:</TH>
				  <TD width="300">&nbsp;<%=shipZipOrPostalCode%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Phone:</TH>
				  <TD width="300">&nbsp;<%=shipPhone%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Email:</TH>
				  <TD width="300">&nbsp;<%=shipEmail%></TD>
				</TR>
				<TR ALIGN="LEFT">
				  <TH align="right" width="200" valign="top">Note:</TH>
				  <TD width="300">&nbsp;<%=shipNote%></TD>
				</TR>
				-->
				  </TABLE>
			<BR>
			
			</center>
<%
		else
			errorMessage = errorMessage & "Error performing transaction."
		end if
	end if

	if( not "" = errorMessage ) then
%>
		<BR><BR><P><center>
		<H2>Error Processing Transaction.</H2>
		<A href="PayPage.html" class = "header"><STRONG>Enter New Payment</STRONG></a>
		</center></p>
		<TABLE CELLPADDING="2" CELLSPACING="2" BORDER="0" width="500" align = "center" ID="Table8">
		<TR class = "header">
		<TD>&nbsp;Error</TD>
		</TR>
		<TR ALIGN="LEFT">
		<TD width="300"><% Response.Write(errorMessage)%></TD>
		</TR>
		</TABLE>
<%
	end if

	'Session.Abandon

	eClient.CleanUp
	set eClient = nothing

	'**********************************************************************************************************

	'Response.End
	
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
	'
		Response.Redirect DCDataDriverExecutedRedirectURL
		'response.Write("<script language='JavaScript'>document.form[0].submit();</script>")
		

	Case "SQLUpdate"
		'...
	Case "SQLMultiDelete"
		'...		
End Select
%>

                            
                            <!--**********************************Transaction result*****************************************-->
                            </td>
                          </tr>
                          <tr> 
                            <td class="CastlesTextBody">&nbsp;</td>
                          </tr>
                        </table>
                      </td>
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                    </tr>
                  </table>
                </td>
                <td bgcolor="#FFFFFF" width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
                <td bgcolor="#CCCCCC" width="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                <td bgcolor="#FFFFFF" width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
                <td width="195" valign="top" bgcolor="#EDEBDB" height="100%"> 
                  <table width="195" border="0" cellspacing="0" cellpadding="0" ID="Table24">
                    <form name="PropertyQuickSearch" method="post" action="subscriptionsuccess.asp" ID="Form2">
                      <tr> 
                        <td width="10" bgcolor="#D6D4BA">&nbsp;</td>
                        <td width="176" bgcolor="#D6D4BA"> 
                          <table width="170" border="0" cellspacing="0" cellpadding="3" height="20" ID="Table25">
                            <tr> 
                              <td class="CastlesTextBodyBold"><%=WebSiteContentCaptionHeader1%></td>
                            </tr>
                          </table>
                        </td>
                        <td width="10" bgcolor="#D6D4BA">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="176" class="CastlesTextBody">&nbsp;</td>
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="176" class="CastlesTextBlack">
                          <table width="175" border="0" cellspacing="0" cellpadding="3" height="20" ID="Table26">
                            <tr> 
                              <td class="CastlesTextBody"><%=WebSiteContentCaption1%></td>
                            </tr>
                          </table>
                        </td>
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="176" class="CastlesTextBodyBold">&nbsp;</td>
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                      </tr>
                    </form>
                  </table>
                </td>
              </tr>
            </table>
            <!-- #EndEditable --></td>
        </tr>
        <tr> 
          <td width="150" valign="top" bgcolor="#FFFFFF" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#FFFFFF" width="196" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td bgcolor="#FFFFFF" width="397" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
        </tr>
        <tr> 
          <td bgcolor="#CCCCCC" width="150" valign="top" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="196" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="396" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td width="150" valign="top" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="196" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="396" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
  	<td width="750">
		<table width="750" border="0" cellspacing="0" cellpadding="0" ID="Table27">
			<tr>
				<td width="150">&nbsp;</td>
				<td width="5">&nbsp;</td>
				
          <td width="595"><!-- #BeginEditable "subBody" --><!-- #EndEditable --></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr> 
    <td valign="top"> 
      <table width="750" border="0" cellspacing="0" cellpadding="0" ID="Table28">
        <tr> 
          <td width="154">&nbsp;</td>
          <td width="595" class="CastlesTextBody">&copy; 2003 <a href="http://www.castlesmag.com" class="normal">Castles 
            Magazine</a> &nbsp;&nbsp;All rights reserved. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="CastlesTextBodyBold">&gt;<a href="../misc/privacypolicy.asp" class="normal">Privacy 
            Policy</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;<a href="../misc/termsofuse.asp" class="normal">Terms 
            of Use</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;
            </span></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
</body>
			
	<!--*************************************************************************************-->		
			
			
				
				
				
				
				
				<!--*************************************************************************************-->
			