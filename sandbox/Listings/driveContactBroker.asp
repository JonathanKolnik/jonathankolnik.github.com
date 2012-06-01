<%@LANGUAGE="VBSCRIPT"%> 

<!--#include virtual="/Templates/castlesdcdatadriver.asp" --> 
<!--#include virtual="/templates/castlesclientcnekt.asp" -->
<%
'Necessary Requests
Function findValue(src)
	dest = request.Form(src)
	if Len(dest)=0 then
		dest = request.QueryString(src)
	end if
	findValue = dest
End Function

SearchStateProvinceID = findValue("SearchStateProvinceID")
SearchCity = findValue("SearchCity")
SearchZipcode = findValue("SearchZipcode")
SearchPriceFrom = findValue("SearchPriceFrom")
SearchPriceTo = findValue("SearchPriceTo")
SearchWaterfront = findValue("SearchWaterfront")
SearchSki = findValue("SearchSki")
SearchCondo = findValue("SearchCondo")
SearchResort = findValue("SearchResort")
SearchCountryClub = findValue("SearchCountryClub")
SearchFarmOrRanch = findValue("SearchFarmOrRanch")

IPAddress = Request.ServerVariables("REMOTE_ADDR")
UserName = CleanForDrive(Request.Form("UserName"))
TelNumber = CleanForDrive(Request.Form("TelNumber"))
EmailAddress = CleanForDrive(Request.Form("EmailAddress"))
Comments = CleanForDrive(Request.Form("Comments"))
ListingID = Request.Form("ListingID")
BrokerID = Request.Form("BrokerID")
ListingAddress = Request.Form("ListingAddress")
ListPrice = Request.Form("ListPrice")
MessageDateTime = now

'get broker's info
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_ClientSide_Broker_Profile"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4) 
	.Parameters.Append .CreateParameter("@BrokerID", 200, 1,200,BrokerID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set BrokerProfile = .Execute()
End With
Set Command1 = Nothing

'Send Broker an email to check his messages.
Message = "Hello " & BrokerProfile.Fields.Item("FirstName").Value & " " & BrokerProfile.Fields.Item("LastName").Value & ", " & vbcrlf & vbcrlf
Message = Message & UserName & " is interested in the following listing: " & vbcrlf & vbcrlf
Message = Message & "Address: " & ListingAddress & vbcrlf
Message = Message & "List Price: " & FormatCurrency(ListPrice,0) & vbcrlf
Message = Message & "ListingID: " & ListingID & vbcrlf & vbcrlf
Message = Message & "The following is the information they entered: " & vbcrlf & vbcrlf
Message = Message & "Name: " & UserName & vbcrlf
Message = Message & "Phone: " & TelNumber & vbcrlf
Message = Message & "Email Address: " & EmailAddress & vbcrlf
Message = Message & "Comments: " & Comments & vbcrlf

Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From = "salesleads@castles.com"
	objCDOMail.Cc = ""
	objCDOMail.Bcc = "jlowenstern@hotmail.com"
	objCDOMail.Subject = "Another Castles Sales Lead for You"
	objCDOMail.To = BrokerProfile.Fields.Item("EmailAddress").Value
	objCDOMail.Body =  Message
	objCDOMail.Send
Set objCDOMail = Nothing

'Insert the message in the brokermessage table.
DataExceptionsString = ""


'Inserts Broker Message
Active="Y"
DCDataDriverType = "SQLInsert"
DBObjectDestination = "Castles_BrokerMessages"
FormType = "Request.Form"
FileServerDestination = ""
FileFate = ""
EmailDestination = ""
DataUniqueKey = "BrokerMessageID"
DataParentNode = ""
DataExceptions = "<!submit!><!ListingAddress!><!ListPrice!>" & DataExceptionsString
DataCookies = ""
DataSessions = ""
DataExtraFields = "IPAddress<!DCDELIMETER!>" & IPAddress & ",MessageDateTime<!DCDELIMETER!>" & MessageDateTime & ",Active<!DCDELIMETER!>" & Active
Cnekt = Connect

BrokerMessageID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)

'Tracks Modification of Broker Messages
EntityID = 13
EntityModificationTypeID = 1
EntityModificationLogger SystemLoginID,EntityID,BrokerMessageID,EntityModificationTypeID,IPAddress 

DCDataDriverExecutedRedirectURL = "ListingDetails.asp?ContactSuccess=Y&ListingID=" & ListingID & "&SearchStateProvinceID=" & Server.URLEncode(SearchStateProvinceID) & "&SearchCity=" & Server.URLEncode(SearchCity) & "&SearchZipcode=" & Server.URLEncode(SearchZipcode) & "&SearchPriceFrom=" & Server.URLEncode(SearchPriceFrom) & "&SearchPriceTo=" & Server.URLEncode(SearchPriceTo) & "&SearchWaterfront=" & Server.URLEncode(SearchWaterfront) & "&SearchSki=" & Server.URLEncode(SearchSki) & "&SearchCondo=" & Server.URLEncode(SearchCondo) & "&SearchResort=" & Server.URLEncode(SearchResort) & "&SearchCountryClub=" & Server.URLEncode(SearchCountryClub) & "&SearchFarmOrRanch=" & Server.URLEncode(SearchFarmOrRanch)

Response.Redirect DCDataDriverExecutedRedirectURL

%>
