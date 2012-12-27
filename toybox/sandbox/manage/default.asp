<%@LANGUAGE="VBSCRIPT"%>
<%
Response.Buffer = True
'Connect = "Provider=sqloledb;Data Source=63.135.108.144;Initial Catalog=castlesmagdb;User Id=castlesclient;Password=45ab#$je;"
'Connect = "Provider=sqloledb;Data Source=192.168.1.204;Initial Catalog=castlesmagdb;User Id=castlesclient;Password=45ab#$je;"
Connect = "Provider=sqloledb;Data Source=mssql06.1and1.com;Initial Catalog=db152651369;User Id=dbo152651369;Password=xYyCaGZG;"

'Color Scheme
TopBar = "#F3F2EA"
NavRegular = "#CF6F19"
NavMuted = "#D6D4BB"
NavDarkBorder = "#666634"
TextBody = "#666634"
TextHeader = "#FFFFFF"
TextLink = "#666634"
TextNavHighlight = "#666634"
TitleBar = "#CF6F19"
LightField = "#F3F2EA"
ThinLine = "#CCCCCC"

Function CleanForDrive(Passenger)
	If Passenger <> "" Then
		CleanForDrive = Replace(Passenger,"'","''")
		CleanForDrive = Trim(CleanForDrive)
	Else
		CleanForDrive = Passenger
	End If
End Function

Function DCSystemLogin(UserName,Password)
	Set Command1 = Server.CreateObject("ADODB.Command")
	With Command1	
		.ActiveConnection = Connect
		.CommandText = "Castles_System_SystemLogin_Validate"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.Parameters.Append .CreateParameter("@UserName", 200, 1,200,UserName)
		.Parameters.Append .CreateParameter("@Password", 200, 1,200,Password)
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		Set SystemLoginValidate = .Execute()
	End With
	Set Command1 = Nothing

	If Not SystemLoginValidate.EOF Then 
		EntityName = SystemLoginValidate.Fields.Item("EntityName").Value 
		SystemLoginID = SystemLoginValidate.Fields.Item("SystemLoginID").Value 
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_SystemLogin_Validate_Info"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@EntityName", 200, 1,200,EntityName)
			.Parameters.Append .CreateParameter("@SystemLoginID", 200, 1,200,SystemLoginID)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set SystemLoginInfo = .Execute()
		End With
		Set Command1 = Nothing
	    
		If Not SystemLoginInfo.EOF Then
			Response.Cookies("SystemLoginID") = SystemLoginID
			Response.Cookies("LanguageID") = SystemLoginValidate.Fields.Item("LanguageID").Value 
			Response.Cookies("CreateNotes") = SystemLoginValidate.Fields.Item("CreateNotes").Value 
			Response.Cookies("DeleteNotes") = SystemLoginValidate.Fields.Item("DeleteNotes").Value 		
			Response.Cookies("SystemLoginNickName") = SystemLoginValidate.Fields.Item("NickName").Value
			Response.Cookies("ShowHelpContent") = SystemLoginValidate.Fields.Item("ShowHelpContent").Value
			Response.Cookies("SystemLoginTypeID") = SystemLoginValidate.Fields.Item("SystemLoginTypeID").Value 
			Response.Cookies("EntityPrimaryKeyValue") = SystemLoginValidate.Fields.Item("EntityPrimaryKeyValue").Value
	
			SystemLoginID = Request.Cookies("SystemLoginID")
			AttemptDateTime = now
			AttemptIP = Request.ServerVariables("REMOTE_ADDR")
			SuccessLogin = "Y"
	
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open Connect
			SQLStmt = "INSERT INTO Castles_SystemLoginLog (AttemptSystemLoginID,AttemptDateTime,AttemptIP,SuccessLogin) "
			SQLStmt = SQLStmt & "VALUES (" & "'" & SystemLoginID & "'" 
			SQLStmt = SQLStmt & "," & "'" & AttemptDateTime & "'" 
			SQLStmt = SQLStmt & "," & "'" & AttemptIP & "'" 
			SQLStmt = SQLStmt & "," & "'" & SuccessLogin & "'" & ")"
			Set SQLAction = Conn.Execute(SQLStmt)
			Set SQLAction = Nothing
			Conn.Close
	
			SystemLoginValidate.Close
			Set SystemLoginValidate = Nothing
			SystemLoginInfo.Close
			Set SystemLoginInfo = Nothing
	
			Response.Redirect "home.asp" 		
			DCSystemLogin = True
		Else
			AttemptDateTime = now
			AttemptIP = Request.ServerVariables("REMOTE_ADDR")
			SuccessLogin = "N"
	
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open Connect
			SQLStmt = "INSERT INTO Castles_SystemLoginLog (AttemptUserName,AttemptPassword,AttemptDateTime,AttemptIP,SuccessLogin) "
			SQLStmt = SQLStmt & "VALUES (" & "'" & UserName & "'" 
			SQLStmt = SQLStmt & "," & "'" & Password & "'" 
			SQLStmt = SQLStmt & "," & "'" & AttemptDateTime & "'" 
			SQLStmt = SQLStmt & "," & "'" & AttemptIP & "'" 
			SQLStmt = SQLStmt & "," & "'" & SuccessLogin & "'" & ")"
			Set SQLAction = Conn.Execute(SQLStmt)
			Set SQLAction = Nothing
			Conn.Close
	
			DCSystemLogin = False
		End if
	Else
		AttemptDateTime = now
		AttemptIP = Request.ServerVariables("REMOTE_ADDR")
		SuccessLogin = "N"

		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Connect
		SQLStmt = "INSERT INTO Castles_SystemLoginLog (AttemptUserName,AttemptPassword,AttemptDateTime,AttemptIP,SuccessLogin) "
		SQLStmt = SQLStmt & "VALUES (" & "'" & UserName & "'" 
		SQLStmt = SQLStmt & "," & "'" & Password & "'" 
		SQLStmt = SQLStmt & "," & "'" & AttemptDateTime & "'" 
		SQLStmt = SQLStmt & "," & "'" & AttemptIP & "'" 
		SQLStmt = SQLStmt & "," & "'" & SuccessLogin & "'" & ")"
		Set SQLAction = Conn.Execute(SQLStmt)
		Set SQLAction = Nothing
		Conn.Close

		DCSystemLogin = False
	End If
End Function

Login = True
UserName = CleanForDrive(Request.Form("UserName"))
Password = CleanForDrive(Request.Form("Password"))
BrokerLogin = Request.QueryString("Broker")

If Len(UserName) <> 0 Then
	Login = DCSystemLogin(UserName,Password)
End If
%>

<html>
<head>
<title>Castles - Management System</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.CastlesTextBlack {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBlackBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBody {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextBody%>; text-decoration: none}
.CastlesTextBodyBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextBody%>; text-decoration: none}
.CastlesTextHeader {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextHeader%>; text-decoration: none}
.CastlesTextHeaderBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextHeader%>; text-decoration: none}
.CastlesTextWhite {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextWhiteBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextNavDark {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextNavHighlight%>; text-decoration: none}

A.navdark:link    { text-decoration: none; color: "<%=TextNavHighlight%>"; }
A.navdark:visited { text-decoration: none; color: "<%=TextNavHighlight%>"; }
A.navdark:active  { text-decoration: none; color: "<%=TextNavHighlight%>"; }
A.navdark:hover   { text-decoration: underline; color: "<%=TextNavHighlight%>"; }

A.white:link    { text-decoration: none; color: "#FFFFFF"; }
A.white:visited { text-decoration: none; color: "#FFFFFF"; }
A.white:active  { text-decoration: none; color: "#FFFFFF"; }
A.white:hover   { text-decoration: underline; color: "#FFFFFF"; }

A.normal:link    { text-decoration: none; color: "<%=TextLink%>"; }
A.normal:visited { text-decoration: none; color: "<%=TextLink%>"; }
A.normal:active  { text-decoration: none; color: "<%=TextLink%>"; }
A.normal:hover   { text-decoration: underline; color: "<%=TextLink%>"; }
-->
</style>
<script language="JavaScript">
<!--
document.onkeypress = KeyHandler;

function KeyHandler(e) {
    if (document.layers){
        Key = e.which;
    }else{
        Key = window.event.keyCode;
		if (Key == 13){
			Login();
			return false;
		}
	}
}

function Login(){
	var ErrorString = ""
	var ErrorTrue = ""
	
	if (document.SystemLogin.UserName.value == "") {
		ErrorString = ErrorString + " - Please enter your user name. \r"
		ErrorTrue="Y"
	}
	if (document.SystemLogin.Password.value == "") {
		ErrorString = ErrorString + " - Please enter your password. \r"
		ErrorTrue="Y"
	}
	if (ErrorTrue == "Y") {
		alert("Missing Required Fields \r" + ErrorString) 
		return false;
	}else {
		document.SystemLogin.submit()
	}
}
function openAccount(){
	opener.location.href="http://www.castlesmag.com/openbrokeraccount/";
	window.close();
}
function forgotPassword(){
	var TheURL = "forgotpassword.asp";
	var WinName = "Password";
	var Features = "width=450,height=300,resizable,scrollbars=yes";
	window.open(TheURL,WinName,Features);
}
function startup(){
	document.SystemLogin.UserName.focus();
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="startup()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="<%=TopBar%>"> 
          <td> 
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="190"><img src="images/castles_logo.GIF" width="190" height="40" usemap="#Map" border="0"></td>
                <td width="410">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="100%" height="20" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="<%=NavDarkBorder%>" height="1"><img src="../Castles/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="../Castles/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavRegular%>" height="19"><img src="../Castles/images/clear10pixel.gif" width="1" height="19"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavDarkBorder%>" height="1"><img src="../Castles/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="../Castles/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="19"> 
            <table width="300" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="CastlesTextNavDark" align=center>&nbsp;<input type=button class="CastlesTextNavDark" name="back" value="Back" onclick="Javascript:window.close();">&nbsp; </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="../Castles/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="10"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
        </tr>
        <tr> 
          <td> 
            <table width="750" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td class="CastlesTextMid" width="3">&nbsp;</td>
                <td class="CastlesTextBodyBig" width="387"><b>
				<%
				if BrokerLogin = "Y" then
					response.write "BROKER "
				end if
				%>
				LOG IN <%If Not(Login) Then%><font color="#990000">*INVALID</font><%End If%></b></td>
                <td class="CastlesTextMid" width="360" align="right">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td height="10"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <form name="SystemLogin" method="post" onSubmit="return Login();" action="default.asp">
                <tr> 
                  <td width="100%" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="100%" bgcolor="<%=LightField%>"> 
                    <table width="750" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td width="200" class="CastlesTextBody">&nbsp;</td>
                        <td width="30" class="CastlesTextBody">&nbsp;</td>
                        <td width="515" class="CastlesTextBody">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td rowspan="5" class="CastlesTextBody" valign="top" width="200"><b>WELCOME 
                          TO THE CASTLES MANAGEMENT SYSTEM</b><br>
                          <br>
                          Please log in to the management system by providing 
                          your username and password in the respective fields 
                          to the right. Submit your login credentials by clicking 
                          on the &quot;Log In&quot; button. If you have forgotten 
                          your password, please click on the &quot;Forgot Your 
                          Password&quot; link and your login credentials will 
                          be emailed to you instantly.</td>
                        <td rowspan="5" class="CastlesTextBody" valign="top" width="30">&nbsp;</td>
                        <td width="515" class="CastlesTextBody">[ 1 ] Username:</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td width="515" class="CastlesTextBlack"> 
                          <input type="text" name="UserName" class="CastlesTextBlack" size="20" maxlength="100">
                        </td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td width="515" class="CastlesTextBody">[ 2 ] Password: 
                        </td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td width="515" class="CastlesTextBlack"> 
                          <input type="password" name="Password" class="CastlesTextBlack" size="17" maxlength="10">
                        </td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td width="515" class="CastlesTextBlack"> 
                          <input type="submit" name="Submit" value="Log In" class="CastlesTextBlack"><br><br>
                          <a href="javascript:forgotPassword()" class="normal"><b>Forgot Your Password?</b></a> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td width="200" class="CastlesTextBody">&nbsp;</td>
                        <td width="30" class="CastlesTextBody">&nbsp;</td>
                        <td width="515" class="CastlesTextBody">
						<%
						if BrokerLogin = "Y" then
						%><b><a href="javascript:openAccount()" class="normal">Click 
                          Here to Open a FREE Broker Account</a></b> - and begin placing your ads online!
						<%
						else
						%>
						&nbsp;
						<%
						end if
						%>
						</td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td width="100%" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <input type="hidden" name="NumberOfRecordsToDelete" value="<%=DeleteCount%>">
              </form>
            </table>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="CastlesTextBody" align="center">&nbsp; </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr width="100%"> 
    <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
  </tr>
  <tr width="100%"> 
    <td width="100%" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td > 
            <table width="481" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="25">&nbsp;</td>
                <td width="456" class="CastlesTextBody">&copy;&nbsp;<%=DatePart("yyyy",Date)%>&nbsp;Castles Magazine.&nbsp;&nbsp;All rights reserved.&nbsp;&nbsp;Powered 
                  By&nbsp;.</td>
              </tr>
            </table>
          </td>
          <td align="right"> 
            <table width="300" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="275" align="right"><a href="http://www.dreamingcode.com" target="_blank"><img src="images/dc_logo_footer.jpg" width="97" height="23" alt="DreamingCode, Inc." border="0"></a></td>
                <td width="25" class="CastlesTextMid">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
