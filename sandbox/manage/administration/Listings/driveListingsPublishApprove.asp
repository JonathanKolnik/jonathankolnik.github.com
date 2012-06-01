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

'Connect = "Provider=sqloledb;Data Source=216.119.113.219;Initial Catalog=castlesdb;User Id=castlesadmin;Password=za7#45g;"
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
NumberOfRecordsToApprove = Request.Form("NumberOfRecordsToApprove")
DataUniqueKey = "ListingID"
count = 1
for i = 1 to NumberOfRecordsToApprove
	if Request.Form("ApproveRecordID" & count) <> "" then
		UniqueKeyValue = Request.Form("ApproveRecordID" & count)
		ListingPublishStatusID = 4
		DesignerPublishStatusID = 1
		DesignerID = 0
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.open Connect
		SQLStmt = "UPDATE Castles_Listings SET ListingPublishStatusID = " & ListingPublishStatusID
		SQLStmt = SQLStmt & ", DesignerPublishStatusID = " & DesignerPublishStatusID
		SQLStmt = SQLStmt & ", DesignerID = " & DesignerID
		SQLStmt = SQLStmt & "WHERE " & DataUniqueKey & " = " & UniqueKeyValue
		Set RS = Conn.Execute(SQLStmt)
		set rs=nothing
		Conn.close

		'Tracks Approval of Records
		EntityModificationTypeID = 7
		EntityModificationLogger SystemLoginID,EntityID,UniqueKeyValue,EntityModificationTypeID,IPAddress 
	end if	
	count = count + 1
next

DCDataDriverExecutedRedirectURL = "publishListings.asp?ListingID=" & Server.URLEncode(SearchListingID) & "&Address=" & Server.URLEncode(SearchedAddress) & "&Sizes=" & Server.URLEncode(SearchedSizes) & "&Areas=" & Server.URLEncode(SearchedAreas) &  "&Status=" & Server.URLEncode(SearchedStatus) & "&PriceFrom=" & Server.URLEncode(PriceFrom) & "&PriceTo=" & Server.URLEncode(PriceTo) & "&Page=" & Page & "&TotalRecords=" & TotalRecords
Response.Redirect DCDataDriverExecutedRedirectURL
%>