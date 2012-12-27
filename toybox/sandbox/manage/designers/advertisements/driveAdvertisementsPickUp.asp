<%@LANGUAGE="VBSCRIPT"%>
<%
'Persist Search Values
Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
SearchListingID = Request.QueryString("SearchListingID")
SearchedAddress = Request.QueryString("SearchedAddress")
SearchedSizes = Request.QueryString("SearchedSizes")
SearchedAreas = Request.QueryString("SearchedAreas")
SearchedStatus = Request.QueryString("SearchedStatus")
PriceFrom = Request.QueryString("PriceFrom")
PriceTo = Request.QueryString("PriceTo")

'Connect = "Provider=sqloledb;Data Source=192.168.1.204;Initial Catalog=castlesdb;User Id=castlesadmin;Password=za7#45g;"
Connect = "Provider=sqloledb;Data Source=mssql06.1and1.com;Initial Catalog=db152651369;User Id=dbo152651369;Password=xYyCaGZG;"

Function EntityModificationLogger(SystemLoginID,EntityID,EntityPrimaryKeyValue,EntityModificationTypeID,IPAddress)
	EntityModificationDateTime = now
	EntityModificationIP = 	IPAddress
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Connect
	SQLStmt = "INSERT INTO Castles_EntityModificationLog (EntityModificationSystemLoginID,EntityID,EntityPrimaryKeyValue,EntityModificationDateTime,EntityModificationIP,EntityModificationTypeID) "
	SQLStmt = SQLStmt & "VALUES (" & "'" & SystemLoginID & "'" 
	SQLStmt = SQLStmt & "," & "'" & EntityID & "'" 
	SQLStmt = SQLStmt & "," & "'" & EntityPrimaryKeyValue & "'" 
	SQLStmt = SQLStmt & "," & "'" & EntityModificationDateTime & "'" 
	SQLStmt = SQLStmt & "," & "'" & EntityModificationIP & "'" 
	SQLStmt = SQLStmt & "," & "'" & EntityModificationTypeID & "'" & ")"
	Set SQLAction = Conn.Execute(SQLStmt)
	Set SQLAction = Nothing
	Conn.Close
End Function


EntityID = 10	
IPAddress = Request.ServerVariables("REMOTE_ADDR")
NumberOfRecordsToPickUp = Request.Form("NumberOfRecordsToPickUp")
DesignerID = Request.Form("DesignerID")
DataUniqueKey = "ListingID"
count = 1
for i = 1 to NumberOfRecordsToPickUp
	'pick-up
	if Request.Form("PickUpRecordID" & count) <> "" then
		UniqueKeyValue = Request.Form("PickUpRecordID" & count)
		DesignerPublishStatusID = 2
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.open Connect
		SQLStmt = "UPDATE Castles_Listings SET DesignerPublishStatusID = " & DesignerPublishStatusID
		SQLStmt = SQLStmt & ", DesignerID = " & DesignerID
		SQLStmt = SQLStmt & "WHERE " & DataUniqueKey & " = " & UniqueKeyValue
		Set RS = Conn.Execute(SQLStmt)
		set rs=nothing
		Conn.close
		'Tracks Approval of Records
		EntityModificationTypeID = 8
		EntityModificationLogger SystemLoginID,EntityID,UniqueKeyValue,EntityModificationTypeID,IPAddress 
	end if
	'drop
	if Request.Form("DropRecordID" & count) <> "" then
		UniqueKeyValue = Request.Form("DropRecordID" & count)
		DesignerPublishStatusID = 1
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.open Connect
		SQLStmt = "UPDATE Castles_Listings SET DesignerPublishStatusID = " & DesignerPublishStatusID
		SQLStmt = SQLStmt & "WHERE " & DataUniqueKey & " = " & UniqueKeyValue
		Set RS = Conn.Execute(SQLStmt)
		set rs=nothing
		Conn.close
		'Tracks Approval of Records
		EntityModificationTypeID = 8
		EntityModificationLogger SystemLoginID,EntityID,UniqueKeyValue,EntityModificationTypeID,IPAddress 
	end if
	'completed
	if Request.Form("CompletedRecordID" & count) <> "" then
		UniqueKeyValue = Request.Form("CompletedRecordID" & count)
		ListingPublishStatusID = 5
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.open Connect
		SQLStmt = "UPDATE Castles_Listings SET ListingPublishStatusID = " & ListingPublishStatusID
		SQLStmt = SQLStmt & "WHERE " & DataUniqueKey & " = " & UniqueKeyValue
		Set RS = Conn.Execute(SQLStmt)
		set rs=nothing
		Conn.close
		'Tracks Approval of Records
		EntityModificationTypeID = 8
		EntityModificationLogger SystemLoginID,EntityID,UniqueKeyValue,EntityModificationTypeID,IPAddress 
	end if	
	count = count + 1
next

DCDataDriverExecutedRedirectURL = "searchAdvertisements.asp?ListingID=" & Server.URLEncode(SearchListingID) & "&Address=" & Server.URLEncode(SearchedAddress) & "&Sizes=" & Server.URLEncode(SearchedSizes) & "&Areas=" & Server.URLEncode(SearchedAreas) &  "&Status=" & Server.URLEncode(SearchedStatus) & "&PriceFrom=" & Server.URLEncode(PriceFrom) & "&PriceTo=" & Server.URLEncode(PriceTo) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
Response.Redirect DCDataDriverExecutedRedirectURL
%>