<%
Response.Buffer = True
'Connect = "Provider=sqloledb;Data Source=63.135.108.144;Initial Catalog=castlesmagdb;User Id=castlesclient;Password=45ab#$je;"
'Connect = "Provider=sqloledb;Data Source=192.168.1.204;Initial Catalog=castlesmagdb;User Id=castlesclient;Password=45ab#$je;"
Connect = "Provider=sqloledb;Data Source=mssql06.1and1.com;Initial Catalog=db152651369;User Id=dbo152651369;Password=xYyCaGZG;"

'Color Scheme
TopBar = "#CBCCD1"
NavRegular = "#C20016"
NavMuted = "#E5997F"
NavDarkBorder = "#660000"
TextBody = "#10344F"
TextHeader = "#3C5A70"
TextLink = "#10344F"
TextNavHighlight = "#660000"
TextNavFade = "#915441"
TitleBar = "#C1CCD1"
LightField = "#E6EBED"

SelfHost = Request.ServerVariables("HOST")
SelfURL = Request.ServerVariables("URL")
SelfQueryString = Request.ServerVariables("QUERY_STRING")
SystemLoginID = Request.Cookies("SystemLoginID")
SystemLoginNickName = Request.Cookies("SystemLoginNickName")
SystemLoginTypeID = Request.Cookies("SystemLoginTypeID")
LanguageID = Request.Cookies("LanguageID")

If LanguageID = "" Then
	LanguageID = 1
End If

'----------------------Begin Global System Functions-----------------------------------------
Function WV(WriteWhatString,WriteWhatValue)
	Response.Write WriteWhatString & " = " & WriteWhatValue & "<br>"
End Function

Function CleanForDrive(Passenger)
	If Passenger <> "" Then
		CleanForDrive = Replace(Passenger,"'","''")
		CleanForDrive = Trim(CleanForDrive)
	Else
		CleanForDrive = Passenger
	End If
End Function

Function DCFormatCurrency(Passenger,RoundToDecimal)
	if Len(Passenger) <> 0 then
		DCFormatCurrency = FormatCurrency(Passenger,RoundToDecimal)
	else
		DCFormatCurrency = FormatCurrency(0,RoundToDecimal)
	end if
End Function

Function FillInSpaceWithNonBreaking(WordPhrase)
	If WordPhrase <> "" Then
		FillInSpaceWithNonBreaking = Replace(WordPhrase," ","&nbsp;")
		FillInSpaceWithNonBreaking = Trim(FillInSpaceWithNonBreaking)
	Else
		FillInSpaceWithNonBreaking = WordPhrase
	End If
End Function

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
'---------------End Global System Functions-------------------------------


%>