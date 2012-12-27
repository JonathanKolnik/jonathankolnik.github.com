<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 1
PageTopNavigationSubHeaderID = 19
FromPageTopNavigationSubHeaderID = 12
EntityID = 5
%>
<!--#include virtual="/templates/castlessystemcnekt.asp" -->
<%
'---------------Begin Page Level Multilingual Translation-----------------------
WordPhrasesOnPage = "Active(|)Company Name(|)Create Notes(|)Delete Notes(|)Designer Info(|)Direct Line(|)Edit Designer(|)Email Address(|)Fax Number(|)First Name(|)Language(|)Last Name(|)Middle Initial(|)Mobile Phone(|)Nick Name(|)No(|)Password(|)Show Help(|)System Access Rights(|)System Notes Access(|)Username(|)View Notes(|)Yes"
WordPhrasesOnPageArray = Split(WordPhrasesOnPage,"(|)")
WhereClause = ""
WordPhraseCount = 1
For Each WordPhrase In WordPhrasesOnPageArray
	If WordPhraseCount = 1 Then
		WhereClause = WhereClause & " EnglishTranslation = '" & WordPhrase & "' AND LanguageID = " & LanguageID 
	Else
		WhereClause = WhereClause & " OR EnglishTranslation = '" & WordPhrase & "' AND LanguageID = " & LanguageID 
	End If
	WordPhraseCount = WordPhraseCount + 1
Next

Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_TranslateWordPhrases_For_Page"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@WhereClause", 201, 1,20000,WhereClause)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set TranslationResults = .Execute()
End With
Set Command1 = Nothing

TranslationResultsArray = TranslationResults.getrows
TranslationResults.close
Set TranslationResults = Nothing
TranslationResultsArrayNumRows = uBound(TranslationResultsArray,2)
TranslateCount = 0
Field_TranslatedWordPhrase = 0

'TestCount = 0
'For Each x In TranslationResultsArray
'WV "x=",TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount+TestCount) & "<br>"
'TestCount = TestCount+1
'Next

WordPhrase_Active = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_CompanyName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_CreateNotes = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_DeleteNotes = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_DesignerInfo = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_DirectLine = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_EditDesigner = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_EmailAddress = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_FaxNumber = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_FirstName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Language = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_LastName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_MiddleInitial = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_MobilePhone = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_NickName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_No = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Password = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_ShowHelp = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SystemAccessRights = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SystemNotesAccess = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Username = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_ViewNotes = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Yes = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)

'---------------End Multilingual Translation-----------------------

'Persist Search Values
Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
SearchColumn = Request.QueryString("SearchColumn")
SearchValue = Request.QueryString("SearchValue")

'Designer Profile Info
DesignerID = Request.QueryString("DesignerID")
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_Designer_Profile_System"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@DesignerID", 200, 1,200,DesignerID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set DesignerProfile = .Execute()
End With
Set Command1 = Nothing

'Designer SystemLogin Info
EntitySystemLoginID = DesignerProfile.Fields.Item("SystemLoginID").Value
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_SystemLogin_Info"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@SystemLoginID", 200, 1,200,EntitySystemLoginID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set DesignerSystemLoginInfo = .Execute()
End With
Set Command1 = Nothing
EntityLanguageID = DesignerSystemLoginInfo.Fields.Item("LanguageID").Value
%>
<html><!-- #BeginTemplate "/Templates/CastlesSystem.dwt" -->
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
.CastlesTextNavFaded {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextNavFaded%>; text-decoration: none}
.CastlesTextNavFadedBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextNavFaded%>; text-decoration: none}

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
<!-- #BeginEditable "script" -->
<script language="JavaScript">
<!--
function Validate(){
	var errorString = ""
	var errorTrue = ""

	if (document.EditEntity.FirstName.value == "") {
		errorString=errorString + " - Please enter your first name. \r"
		errorTrue="y"
	}
	if (document.EditEntity.LastName.value == "") {
		errorString=errorString + " - Please enter your last name. \r"
		errorTrue="y"
	}
	if (document.EditEntity.EmailAddress.value == "") {
		errorString=errorString + " - Please enter your email address. \r"
		errorTrue="y"
	}
	if (document.EditEntity.UserName.value == "") {
		errorString=errorString + " - Please enter your username. \r"
		errorTrue="y"
	}
	if (document.EditEntity.Password.value == "") {
		errorString=errorString + " - Please enter your password. \r"
		errorTrue="y"
	}

	if (errorTrue == "y") {
		alert("Missing Required Fields: \r" + errorString) 
		return false;
	}else {
		return true;
	}
}
//-->
</script>
<!-- #EndEditable --> 
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="<%=TopBar%>"> 
          <td> 
            <table width="750" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="190"><img src="../../../images/castles_logo.GIF" width="190" height="40" usemap="#Map" border="0"></td>
                <td width="560"> 
                  <table width="560" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="CastlesTextBody" width="230" align="right"><%=FillInSpaceWithNonBreaking(WordPhrase_Welcome)%>,&nbsp;<%=SystemLoginNickName%>&nbsp;&nbsp;<b>&gt;&nbsp;<a href="../../../editProfile.asp" class="normal"><%=FillInSpaceWithNonBreaking(WordPhrase_EditProfile)%></a></b>&nbsp;&nbsp;<b>&gt;&nbsp;<a href="../../../logout.asp" class="normal"><%=FillInSpaceWithNonBreaking(WordPhrase_LogOut)%></a>&nbsp;&nbsp;&nbsp;</b></td>
                      <td width="330" align="right"> 
                        <table width="300" border="0" cellspacing="0" cellpadding="0">
                          <form name="QuickTaskRedirect" method="post" action="/manage/quicktaskredirect.asp">
                            <tr> 
                              <td class="CastlesTextBody" align="right" width="60"><%=FillInSpaceWithNonBreaking(WordPhrase_QuickTask)%>:&nbsp;</td>
                              <td class="CastlesTextBody" width="240"> 
<%
QuickTaskURL = SelfHost & SelfURL & "?" & SelfQueryString
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_QuickTask_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@SystemLoginID", 200, 1,200,SystemLoginID)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,LanguageID)	
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
Set QuickTaskList = .Execute()
End With
Set Command1 = Nothing

QuickTaskListCode = ""
While Not QuickTaskList.EOF
	RedirectToPath = QuickTaskList.Fields.Item("RedirectToPath").Value
	TopNavigationSubHeaderName = QuickTaskList.Fields.Item("TopNavigationSubHeaderName").Value
	QuickTaskListCode = QuickTaskListCode & "<option value=""" & RedirectToPath & """>" & TopNavigationSubHeaderName & "</option>"
	QuickTaskList.MoveNext()
Wend
%>

                                <select name="QuickTaskURL" class="CastlesTextBlack" onChange="document.QuickTaskRedirect.submit();">
                                  <option value="<%=QuickTaskURL%>" selected><%=WordPhrase_SelectAQuickTask%></option>
								  <%=QuickTaskListCode%>
                                </select>&nbsp;&nbsp;<input type="submit" name="Submit" value="<%=WordPhrase_Go%>" class="CastlesTextBlack">
                              </td>
                            </tr>
                          </form>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="100%" height="20" valign="top"><%=TopNavCode%></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <table width="750" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td class="CastlesTextMid" colspan="3" height="1"><img src="file:///C|/Clients/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td width="3">&nbsp;</td>
          <td class="CastlesTextNavFadedBig" width="587"><%=PageBreadCrumb%></td>
<%
'Shows or Hides Help Content And Generates Proper Link to Change
FormElementCount = 1
For Each DataElement In Request.Form 
	ElementName =  DataElement 
	ElementValue = CleanForDrive(Request.Form(ElementName))
	If FormElementCount = 1 Then
		If Len(SelfQueryString) = 0 Then
			SelfQueryStringStart = "?"
		Else
			SelfQueryStringStart = "&"
		End If
		SelfQueryString = SelfQueryString & SelfQueryStringStart & ElementName & "=" & ElementValue
	Else
		SelfQueryString = SelfQueryString & "&" & ElementName & "=" & ElementValue
	End If
	FormElementCount = FormElementCount + 1
Next 			


If Len(SelfQueryString) <> 0 Then
	SelfQueryString = Replace(SelfQueryString,"ShowHelpContent=Y","")
	SelfQueryString = Replace(SelfQueryString,"HideHelpContent=Y","")
End If

SelfLink = SelfHost & SelfURL & "?" & SelfQueryString
If inStr(SelfLink,"?") = 0 Then
	SelfLink = SelfLink & "?"
Else
	SelfLink = SelfLink & "&"
End If

SelfLink = Replace(SelfLink,"?&","?")
SelfLink = Replace(SelfLink,"&&","&")
SelfLink = Replace(SelfLink,"??","?")

If SystemHelpContentText = "" Then
	SelfLink = SelfLink & "ShowHelpContent=Y"
%>
          <td class="CastlesTextBody" width="160" align="right"><a href="<%=SelfLink%>" class="normal"><b><%=WordPhrase_ShowHelpText%></b></a></td>
<%
Else
	SelfLink = SelfLink & "HideHelpContent=Y"
%>
          <td class="CastlesTextBody" width="360" align="right"><a href="<%=SelfLink%>" class="normal"><b><%=WordPhrase_HideHelpText%></b></a></td>
<%
End If
%>
        </tr>
<%
If SystemHelpContentText <> "" Then
%>
        <tr> 
          <td width="3">&nbsp;</td>
          <td class="CastlesTextBody" colspan="2"><%=SystemHelpContentText%></td>
        </tr>
<%
End If
%>
        <tr> 
          <td class="CastlesTextBody" colspan="3" height="1"><img src="file:///C|/Clients/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"><!-- #BeginEditable "body" -->
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="<%=ThinLine%>"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#000000" width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td bgcolor="<%=TitleBar%>" width="100%" class="CastlesTextHeader" height="20">&nbsp;&nbsp;&nbsp;<b><%=UCase(WordPhrase_DesignerInfo)%></b></td>
              </tr>
              <form name="EditEntity" method="post" onSubmit="return Validate();" action="drivedesigners.asp?DCDataDriverType=SQLUpdate&SearchColumn=<%=SearchColumn%>&SearchValue=<%=SearchValue%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>">
                <tr> 
                  <td width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="100%" bgcolor="<%=LightField%>"> 
                    <table width="750" border="0" cellspacing="0" cellpadding="2">
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td rowspan="14" valign="top"> 
                          <table width="350" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;1&nbsp;] <%=WordPhrase_FirstName%> / <%=WordPhrase_MiddleInitial%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="FirstName" class="CastlesTextBlack" size="15" maxlength="50" value="<%=DesignerProfile.Fields.Item("FirstName").Value%>">
                                &nbsp;&nbsp; 
                                <input type="text" name="MiddleInitial" class="CastlesTextBlack" size="3" maxlength="1" value="<%=DesignerProfile.Fields.Item("MiddleInitial").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;2&nbsp;] <%=WordPhrase_LastName%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="LastName" class="CastlesTextBlack" size="20" maxlength="50" value="<%=DesignerProfile.Fields.Item("LastName").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;3&nbsp;] <%=WordPhrase_CompanyName%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="CompanyName" class="CastlesTextBlack" size="20" maxlength="40" value="<%=DesignerProfile.Fields.Item("CompanyName").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;4&nbsp;] <%=WordPhrase_DirectLine%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="DirectLine" class="CastlesTextBlack" size="20" maxlength="40" value="<%=DesignerProfile.Fields.Item("DirectLine").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;5&nbsp;] <%=WordPhrase_MobilePhone%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="MobileTelNumber" class="CastlesTextBlack" size="20" maxlength="40" value="<%=DesignerProfile.Fields.Item("MobileTelNumber").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;6&nbsp;] <%=WordPhrase_FaxNumber%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="FaxNumber" class="CastlesTextBlack" size="20" maxlength="40" value="<%=DesignerProfile.Fields.Item("FaxNumber").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;7&nbsp;] <%=WordPhrase_EmailAddress%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="EmailAddress" class="CastlesTextBlack" size="30" maxlength="100" value="<%=DesignerProfile.Fields.Item("EmailAddress").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;8&nbsp;] <%=WordPhrase_Language%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <%
LanguageList = ""
ActiveLanguageList = ActiveLanguages()
LanguageListArray = Split(ActiveLanguageList,",")
LanguagesCount = 1

For Each Language In LanguageListArray
	LanguageSpecificArray = Split(Language,"<!DCDELIMETER!>")
	SpecificLanguageName = LanguageSpecificArray(0)
	SpecificLanguageID = LanguageSpecificArray(1)

	If cStr(SpecificLanguageID) = cStr(EntityLanguageID) Then
		LanguageList = LanguageList & "<option value=""" & SpecificLanguageID & """ selected>" & SpecificLanguageName & "</option>" & vbcrlf
	Else
		LanguageList = LanguageList & "<option value=""" & SpecificLanguageID & """>" & SpecificLanguageName & "</option>" & vbcrlf
	End If
	LanguagesCount = LanguagesCount + 1
Next
%>
                                <select name="EntityLanguageID" class="CastlesTextBlack">
                                  <%=LanguageList%> 
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;9&nbsp;] <%=WordPhrase_Active%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="Active" class="CastlesTextBlack">
                                  <%If DesignerProfile.Fields.Item("Active").Value = "Y" Then%>
                                  <option value="Y" selected><%=WordPhrase_Yes%></option>
                                  <option value="N"><%=WordPhrase_No%></option>
                                  <%Else%>
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                  <%End If%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;10&nbsp;] <%=WordPhrase_NickName%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="NickName" class="CastlesTextBlack" size="20" maxlength="10" value="<%=DesignerSystemLoginInfo.Fields.Item("NickName").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;11&nbsp;] <%=WordPhrase_UserName%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="UserName" class="CastlesTextBlack" size="20" maxlength="100" value="<%=DesignerSystemLoginInfo.Fields.Item("UserName").Value%>">
                              </td>
                            </tr>
                            <%
If  Request.QueryString("DuplicateFound") = "Y" Then                      
%>
                            <tr> 
                              <td class="CastlesTextBody" valign="top"> <font color="#993333">* 
                                The username &quot;<%=Request.QueryString("AttemptUserName")%>&quot; has already been taken! Please 
                                choose another or use your email address. </font></td>
                            </tr>
                            <%
End If
%>
                            <tr> 
                              <td class="CastlesTextBody"> [&nbsp;12&nbsp;] <%=WordPhrase_Password%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="password" name="Password" class="CastlesTextBlack" size="17" maxlength="10"  value="<%=DesignerSystemLoginInfo.Fields.Item("Password").Value%>">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;13&nbsp;] <%=WordPhrase_ShowHelp%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="ShowHelpContent" class="CastlesTextBlack">
                                  <%If DesignerSystemLoginInfo.Fields.Item("ShowHelpContent").Value = "Y" Then%>
                                  <option value="Y" selected><%=WordPhrase_Yes%></option>
                                  <option value="N"><%=WordPhrase_No%></option>
                                  <%Else%>
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                  <%End If%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;14&nbsp;] <%=WordPhrase_SystemNotesAccess%>:</td>
                            </tr>
                            <tr> 
                              <td> 
                                <table width="200" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td width="20"> 
                                      <%
If DesignerSystemLoginInfo.Fields.Item("CreateNotes").Value = "Y" Then
%>
                                      <input type="checkbox" name="CreateNotes" value="Y" checked>
                                      <%
Else
%>
                                      <input type="checkbox" name="CreateNotes" value="Y">
                                      <%
End If
%>
                                    </td>
                                    <td width="180" class="CastlesTextBody"><%=WordPhrase_CreateNotes%> </td>
                                  </tr>
                                  <tr> 
                                    <td width="20"> 
                                      <%
If DesignerSystemLoginInfo.Fields.Item("DeleteNotes").Value = "Y" Then
%>
                                      <input type="checkbox" name="DeleteNotes" value="Y" checked>
                                      <%
Else
%>
                                      <input type="checkbox" name="DeleteNotes" value="Y">
                                      <%
End If
%>
                                    </td>
                                    <td width="180" class="CastlesTextBody"><%=WordPhrase_DeleteNotes%> </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                        </td>
                        <td width="365" class="CastlesTextBody">[&nbsp;15&nbsp;] 
                          <%=WordPhrase_SystemAccessRights%>:</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td rowspan="13" valign="top"> 
                          <table width="280" border="0" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
                            <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_SystemLoginAccessRights_Profile"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@SystemLoginID", 200, 1,200,EntitySystemLoginID)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,EntityLanguageID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set SystemLoginAccessRightsProfile = .Execute()
End With
Set Command1 = Nothing

HeaderAccessRightsString = ""
SubHeaderAccessRightsString = ""
While Not SystemLoginAccessRightsProfile.EOF
	TopNavigationHeaderID = SystemLoginAccessRightsProfile.Fields.Item("TopNavigationHeaderID").Value
	TopNavigationSubHeaderID = SystemLoginAccessRightsProfile.Fields.Item("TopNavigationSubHeaderID").Value

	If TopNavigationSubHeaderID = 0 Then
		HeaderAccessRightsString = HeaderAccessRightsString & "(" & TopNavigationHeaderID & ")"
	Else
		SubHeaderAccessRightsString = SubHeaderAccessRightsString & "(" & TopNavigationSubHeaderID & ")"
	End If
	SystemLoginAccessRightsProfile.MoveNext()
Wend

SystemLoginTypeID = DesignerSystemLoginInfo.Fields.Item("SystemLoginTypeID").Value
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_SystemLoginAccessRights_List_System"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@SystemLoginTypeID", 200, 1,200,SystemLoginTypeID)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,LanguageID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set SystemLoginAccessRightsList = .Execute()
End With
Set Command1 = Nothing

HeaderCount = 1
SubHeaderCount = 1
LastTopNavigationHeaderName = ""
While Not SystemLoginAccessRightsList.EOF 
	TopNavigationHeaderID = SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderID").Value
	ParenTopNavigationHeaderID = "(" & TopNavigationHeaderID & ")"
	TopNavigationSubHeaderID = SystemLoginAccessRightsList.Fields.Item("TopNavigationSubHeaderID").Value
	ParenTopNavigationSubHeaderID = "(" & TopNavigationSubHeaderID & ")"

	If LastTopNavigationHeaderName <> SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderName").Value Then
%>
                            <tr> 
                              <td width="15" class="CastlesTextHeader" bgcolor="<%=DarkColor%>"> 
                                <%
		If inStr(HeaderAccessRightsString,ParenTopNavigationHeaderID) = 0 Then
%>
                                <input type="checkbox" name="TopNavigationHeaderID<%=HeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderID").Value%>">
                                <%
		Else
%>
                                <input type="checkbox" name="TopNavigationHeaderID<%=HeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderID").Value%>" checked>
                                <%
		End If
%>
                              </td>
                              <td width="265" colspan="2" class="CastlesTextBody" bgcolor="<%=DarkColor%>"><b><%=SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderName").Value%></b></td>
                            </tr>
                            <tr> 
                              <td width="15" bgcolor="<%=LightField%>">&nbsp;</td>
                              <td width="15" class="CastlesTextBody" bgcolor="<%=LightColor%>"> 
                                <input type="hidden" name="TopNavigationHeaderIDForSubHeader<%=SubHeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderID").Value%>">
                                <%
		If inStr(SubHeaderAccessRightsString,ParenTopNavigationSubHeaderID) = 0 Then
%>
                                <input type="checkbox" name="TopNavigationSubHeaderID<%=SubHeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationSubHeaderID").Value%>">
                                <%
		Else
%>
                                <input type="checkbox" name="TopNavigationSubHeaderID<%=SubHeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationSubHeaderID").Value%>" checked>
                                <%
		End If
%>
                              </td>
                              <td width="250" class="CastlesTextBody" bgcolor="<%=LightField%>"><%=SystemLoginAccessRightsList.Fields.Item("TopNavigationSubHeaderName").Value%></td>
                            </tr>
                            <%
		HeaderCount = HeaderCount + 1
		SubHeaderCount = SubHeaderCount + 1
	Else
%>
                            <tr> 
                              <td width="15" bgcolor="<%=LightField%>">&nbsp;</td>
                              <td width="15" class="CastlesTextBody" bgcolor="<%=LightField%>"> 
                                <input type="hidden" name="TopNavigationHeaderIDForSubHeader<%=SubHeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderID").Value%>">
                                <%
		If inStr(SubHeaderAccessRightsString,ParenTopNavigationSubHeaderID) = 0 Then
%>
                                <input type="checkbox" name="TopNavigationSubHeaderID<%=SubHeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationSubHeaderID").Value%>">
                                <%
		Else
%>
                                <input type="checkbox" name="TopNavigationSubHeaderID<%=SubHeaderCount%>" value="<%=SystemLoginAccessRightsList.Fields.Item("TopNavigationSubHeaderID").Value%>" checked>
                                <%
		End If
%>
                              </td>
                              <td width="250" class="CastlesTextBody" bgcolor="<%=LightField%>"><%=SystemLoginAccessRightsList.Fields.Item("TopNavigationSubHeaderName").Value%></td>
                            </tr>
                            <%
		SubHeaderCount = SubHeaderCount + 1
	End If
	LastTopNavigationHeaderName = SystemLoginAccessRightsList.Fields.Item("TopNavigationHeaderName").Value
	SystemLoginAccessRightsList.MoveNext()
Wend
SystemLoginAccessRightsList.Close
Set SystemLoginAccessRightsList = Nothing
%>
                          </table>
                          <br>
                          <table width="300" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="hidden" name="DesignerID" value="<%=DesignerID%>">
                                <input type="hidden" name="EntitySystemLoginID" value="<%=EntitySystemLoginID%>">
                                <input type="hidden" name="HeaderCount" value="<%=HeaderCount%>">
                                <input type="hidden" name="SubHeaderCount" value="<%=SubHeaderCount%>">
                                <input type="submit" name="Submit" value="<%=WordPhrase_EditDesigner%>" class="CastlesTextBlack">
                              </td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
              </form>
            </table>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td align="center">&nbsp; </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <!-- #EndEditable --></td>
  </tr>
  <tr width="100%"> 
    <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="../../../images/clear10pixel.gif" width="1" height="1"></td>
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
                <td width="275" align="right"><a href="http://www.dreamingcode.com" target="_blank"><img src="/manage/images/dc_logo_footer.jpg" width="97" height="23" alt="DreamingCode, Inc." border="0"></a></td>
                <td width="25" class="CastlesTextMid">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<map name="Map"> 
  <area shape="rect" coords="16,5,161,38" href="/manage/home.asp" alt="Home" title="Home">
</map>
</body>
<!-- #EndTemplate --></html>
