<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/templates/castlesdcdatadriver.asp" --> 
<%
'Persist Search Values
Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
SearchListingID = Request.QueryString("SearchListingID")
SearchedAddress = Request.QueryString("SearchedAddress")
SearchedSizes = Request.QueryString("SearchedSizes")
SearchedStates = Request.QueryString("SearchedStates")
SearchedApartmentStatus = Request.QueryString("SearchedApartmentStatus")
PriceFrom = Request.QueryString("PriceFrom")
PriceTo = Request.QueryString("PriceTo")

'Which DCDataDriverType to use
DCDataDriverType = Request.QueryString("DCDataDriverType")
if DCDataDriverType <> "SQLMultiDelete" then
	'Necessary Requests
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	upl.Path = Server.Mappath ("/manage/images")
	
	ListingID = upl.Form("ListingID")
	ListPrice = CleanForDriveFloat(upl.Form("ListPrice"))
	ShowAddress= CleanForDriveFloat(upl.Form("ShowAddress"))
	ShowListPrice= upl.Form("ShowListPrice")
	Assessment = CleanForDriveFloat(upl.Form("Assessment"))
	Taxes = CleanForDriveFloat(upl.Form("Taxes"))
	
	'PictureWidth1 = upl.form("PictureWidth1")
	'PictureHeight1 = upl.form("PictureHeight1")
	'PicturePath1 = upl.form("PicturePath1")
	'PictureWidth2 = upl.form("PictureWidth2")
	'PictureHeight2 = upl.form("PictureHeight2")
	'PicturePath2 = upl.form("PicturePath2")
	
	for i = 1 to 8
		DataExceptionsString = DataExceptionsString & "<!PictureWidth" & i & "!>"
		DataExceptionsString = DataExceptionsString & "<!PictureHeight" & i & "!>"
		DataExceptionsString = DataExceptionsString & "<!PicturePath" & i & "!>"
	next
end if

if isNull(ShowAddress) or IsEmpty(ShowAddress) or ShowAddress = "" then
	ShowAddress = "N"
else 
	ShowAddress = "Y"				
end if

if isNull(ShowListPrice) or IsEmpty(ShowListPrice) or ShowListPrice = "" then
	ShowListPrice = "N"
else 
	ShowListPrice = "Y"
end if
ShowListPrice = "Y"

'Handle different DCDataDriverTypes
Select Case DCDataDriverType
	Case "SQLInsert"
		'Inserts Listing Information
		DBObjectDestination = "Castles_Listings"
		FormType = "upl.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "ListingID"
		DataParentNode = ""
		DataExceptions = "<!ListingID!><!ListPrice!><!Assessment!><!Taxes!><!Submit!><!ListingPublishStatusID!>" & DataExceptionsString & "<!ShowAddress!><!ShowListPrice!>"
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = "ListPrice<!DCDELIMETER!>" & ListPrice & ",Assessment<!DCDELIMETER!>" & Assessment & ",Taxes<!DCDELIMETER!>" & Taxes & ",ListingPublishStatusID<!DCDELIMETER!>" & 0 & ",ShowAddress<!DCDELIMETER!>" & ShowAddress & ",ShowListPrice<!DCDELIMETER!>" & ShowListPrice
		Cnekt = Connect
		
		ListingID = DCDataDriver(DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType)
		
		'Tracks Modification of Listing
		EntityID = 10
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,BrokerID,EntityModificationTypeID,IPAddress 

		DCDataDriverExecutedRedirectURL = "searchListings.asp?ListingID=" & Server.URLEncode(ListingID) & "&SearchedAddress=" & Server.URLEncode(SearchedAddress) & "&SearchedSizes=" & Server.URLEncode(SearchedSizes) & "&SearchedStates=" & Server.URLEncode(SearchedStates) & "&SearchedApartmentStatus=" & Server.URLEncode(SearchedApartmentStatus) & "&PriceFrom=" & Server.URLEncode(PriceFrom) & "&PriceTo=" & Server.URLEncode(PriceTo) & "&Page=" & Page & "&TotalRecords=" & TotalRecords

		'update database for the picture
		count = 1
		for each item in upl.form
			If IsObject(upl.form(item)) then
				If upl.form(item).IsEmpty Then
				Else	   		
					PicturePath = "http://castlesmag.com/manage/images/Listing" & ListingID & "_" & count & ".jpg"
					NewFileName = "Listing" & ListingID & "_" & count & ".jpg"
					upl.Form(item).SaveAs NewFileName
					'upamadate the datamabase
					Set Conn = Server.CreateObject("ADODB.Connection")
					Conn.open connect
					SQLStmt = "UPDATE Castles_Listings SET "
					SQLStmt = SQLStmt & "PicturePath" & count & " ='" & PicturePath & "' "
					SQLStmt = SQLStmt & "WHERE ListingID ='" & ListingID & "'"
					Set RS = Conn.Execute(SQLStmt)
					set rs=nothing
					Conn.close
				End If
				count = count + 1
			end if
		next
		
		'Code to update showaddress
			'Set Conn = Server.CreateObject("ADODB.Connection")
			'Conn.open connect
			
			'Response.Write("ShowAddress = " & ShowAddress)  
			'if isNull(ShowAddress) or IsEmpty(ShowAddress) or ShowAddress = "" then
				'ShowAddress = 'N'
				'SQLStmt =  "UPDATE Castles_Listings set ShowAddress = 'N' WHERE ListingID =" & ListingID
			'else 
				'ShowAddress = 'Y'
				'SQLStmt =  "UPDATE Castles_Listings set ShowAddress = 'Y' WHERE ListingID =" & ListingID
			'end if
			'Response.Write(SQLStmt)
			'Response.End()
			'Set RS = Conn.Execute(SQLStmt)
			'set rs=nothing
			'Conn.close
			
			'Code to update showlistprice
			'Set Conn = Server.CreateObject("ADODB.Connection")
			'Conn.open connect
			
			'if isNull(ShowListPrice) or IsEmpty(ShowListPrice) or ShowListPrice = "" then
				'ShowListPrice = 'N'
				'SQLStmt =  "UPDATE Castles_Listings set ShowListPrice = 'N' WHERE ListingID =" & ListingID
			'else 
				'ShowListPrice = 'Y'
				'SQLStmt =  "UPDATE Castles_Listings set ShowListPrice = 'Y' WHERE ListingID =" & ListingID
			'end if
			'Response.Write(SQLStmt)
			'Response.End()
			'Set RS = Conn.Execute(SQLStmt)
			'set rs=nothing
			'Conn.close
		
		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLUpdate"
		'Updates Broker Information		
		DBObjectDestination = "Castles_Listings"
		FormType = "upl.Form"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "ListingID"
		DataParentNode = ""
		DataExceptions = "<!ListingID!><!ListPrice!><!ShowAddress!><!ShowListPrice!><!Assessment!><!Taxes!><!Submit!><!ListingPublishStatusID!>" & DataExceptionsString
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = "ListPrice<!DCDELIMETER!>" & ListPrice & ",Assessment<!DCDELIMETER!>" & Assessment & ",Taxes<!DCDELIMETER!>" & Taxes & ",ShowAddress<!DCDELIMETER!>" & ShowAddress & ",ShowListPrice<!DCDELIMETER!>" & ShowListPrice 
		
		DCDataDriverExecutedRedirectURL = "searchListings.asp?SearchListingID=" & Server.URLEncode(SearchListingID) & "&SearchedAddress=" & Server.URLEncode(SearchedAddress) & "&SearchedSizes=" & Server.URLEncode(SearchedSizes) & "&SearchedStates=" & Server.URLEncode(SearchedStates) & "&SearchedApartmentStatus=" & Server.URLEncode(SearchedApartmentStatus) & "&PriceFrom=" & Server.URLEncode(PriceFrom) & "&PriceTo=" & Server.URLEncode(PriceTo) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		
		Cnekt = Connect
		
		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		'Tracks Modification of Listing
		EntityID = 10
		EntityModificationTypeID = 2
		EntityModificationLogger SystemLoginID,EntityID,BrokerID,EntityModificationTypeID,IPAddress 
		
		'update database for the picture
		count = 1
		for each item in upl.form
			If IsObject(upl.form(item)) then
				If upl.form(item).IsEmpty Then
				Else	   		
					PicturePath = "http://castlesmag.com/manage/images/Listing" & ListingID & "_" & count & ".jpg"
					NewFileName = "Listing" & ListingID & "_" & count & ".jpg"
					upl.Form(item).SaveAs NewFileName
					'upamadate the datamabase
					Set Conn = Server.CreateObject("ADODB.Connection")
					Conn.open connect
					SQLStmt = "UPDATE Castles_Listings SET "
					SQLStmt = SQLStmt & "PicturePath" & count & " ='" & PicturePath & "', "
					SQLStmt = SQLStmt & "PictureWidth" & count & " ='" & upl.form("PictureWidth"&count) & "', "
					SQLStmt = SQLStmt & "PictureHeight" & count & " ='" & upl.form("PictureHeight"&count) & "' "
					SQLStmt = SQLStmt & "WHERE ListingID ='" & ListingID & "'"
					
					Set RS = Conn.Execute(SQLStmt)
					set rs=nothing
					Conn.close
					
				End If
				count = count + 1
			end if			 
		next
		
		Response.Redirect DCDataDriverExecutedRedirectURL

	Case "SQLMultiDelete"
		'Deletes Broker Information		
		DBObjectDestination = "Castles_Listings"
		FileServerDestination = ""
		FileFate = ""
		EmailDestination = ""
		DataUniqueKey = "ListingID"
		DataParentNode = ""
		DataExceptions = ""
		DataCookies = ""
		DataSessions = ""
		DataExtraFields = ""
		DCDataDriverExecutedRedirectURL = "searchListings.asp?SearchListingID=" & Server.URLEncode(SearchListingID) & "&SearchedAddress=" & Server.URLEncode(SearchedAddress) & "&SearchedSizes=" & Server.URLEncode(SearchedSizes) & "&SearchedStates=" & Server.URLEncode(SearchedStates) & "&SearchedApartmentStatus=" & Server.URLEncode(SearchedApartmentStatus) & "&PriceFrom=" & Server.URLEncode(PriceFrom) & "&PriceTo=" & Server.URLEncode(PriceTo) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
		Cnekt = Connect
		EntityID = 10

		DCDataDriver DCDataDriverType,DBObjectDestination,FileServerDestination,FileFate,EmailDestination,DataUniqueKey,DataParentNode,DataCookies,DataSessions,DataExtraFields,DataExceptions,Cnekt,EntityID,FormType
		
		Response.Redirect DCDataDriverExecutedRedirectURL
End Select

%>