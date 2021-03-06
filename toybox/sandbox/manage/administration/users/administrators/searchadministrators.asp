<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 1
PageTopNavigationSubHeaderID = 1
EntityID = 1
%>
<!--#include virtual="/templates/castlessystemcnekt.asp" -->
<!--#include virtual="/templates/castlesdcsystemsimplesearch.asp" --> 
<%
'---------------Begin Page Level Multilingual Translation-----------------------
WordPhrasesOnPage = "Actions(|)Active(|)Administrator(|)Delete(|)Direct Line(|)Edit(|)Email Address(|)Jump To Page(|)Mobile Phone(|)Name(|)No Results Found(|)Notes(|)Search(|)Search By(|)Search Field(|)Search Value(|)View"
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

WordPhrase_Actions = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Active = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Administrator = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Delete = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_DirectLine = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Edit = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_EmailAddress = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_JumpToPage = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_MobilePhone = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Name = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_NoResultsFound = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Notes = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Search = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchBy = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchField = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchValue = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_View = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'---------------End Multilingual Translation-----------------------


RecsPerPage = 15

Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
MoreRecords = Request.QueryString("MoreRecords")
LastSearchValue = Request.Form("LastSearchValue")
SearchOnEntity = "Administrators"

If Len(Request.Form("SearchColumn")) <> 0 Then
	SearchColumn = Request.Form("SearchColumn")
	SearchValue = Request.Form("SearchValue")
Else
	SearchColumn = Request.QueryString("SearchColumn")
	SearchValue = Request.QueryString("SearchValue")
End If

If Len(SearchColumn) <> 0 Then
	SearchResultsReturned = DCSystemSimpleSearchHeader(Page,RecsPerPage,SearchOnEntity,SearchColumn,SearchValue)
	If Len(SearchValue) <> 0 Then
		SearchValue = Replace(SearchValue,"''","'")
	End If
End If

'Allows Entity Modification Access By SystemLoginID
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_EntityModification_AccessRights"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@SystemLoginID", 200, 1,200,SystemLoginID)
	.Parameters.Append .CreateParameter("@TopNavigationHeaderID", 200, 1,200,PageTopNavigationHeaderID)
	.Parameters.Append .CreateParameter("@EntityID", 200, 1,200,EntityID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set EntityModificationAccessRights = .Execute()
End With
Set Command1 = Nothing

EntityModificationAccessString = ""

While Not EntityModificationAccessRights.EOF
		EntityModificationTypeID = EntityModificationAccessRights.Fields.Item("EntityModificationTypeID").Value
		EntityModificationAccessString = EntityModificationAccessString & "(" & EntityModificationTypeID & ")"
	EntityModificationAccessRights.MoveNext()
Wend

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
function PopNotes(EntityID,EntityPrimaryKeyValue,PersonalName){
	EntityName = "<%=WordPhrase_Administrator%>"
	TheURL = "/manage/administration/notes/noteslist.asp?EntityID="+EntityID+"&EntityPrimaryKeyValue="+EntityPrimaryKeyValue+"&EntityName="+EntityName+"&PersonalName="+PersonalName+"&PageTopNavigationSubHeaderID=10"
	WinName = "NotesList"+EntityPrimaryKeyValue
	Features = "width=720,height=608,resizable,scrollbars=yes"
	window.open(TheURL,WinName,Features);
}  


function DeleteCheckedRecords(){
	if(confirm("Are you sure you want to permanently delete the selected item(s)?")){
		return true;
	}else {
		return false;
	}
}

function Search(){
	var ErrorString = ""
	var ErrorTrue = ""
	
	if (document.SearchRecords.SearchValue.value == "") {
		ErrorString = ErrorString + "Please enter a value for the search. \r"
		ErrorTrue="Y"
	}
	if (ErrorTrue == "Y") {
		alert(ErrorString) 
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
          <td> 
            <table width="750" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td align="right" colspan="4" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <form name="SearchRecords" method="post" onSubmit="return Search();" action="searchadministrators.asp">
                <tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b><%=FillInSpaceWithNonBreaking(WordPhrase_SearchBy)%>:</b>&nbsp;</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;1&nbsp;]&nbsp;<%=FillInSpaceWithNonBreaking(WordPhrase_SearchField)%>:</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;2&nbsp;]&nbsp;<%=FillInSpaceWithNonBreaking(WordPhrase_SearchValue)%>:</td>
                  <td class="CastlesTextBody" width="385">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="65">&nbsp;</td>
                  <td class="CastlesTextBody" width="150"> 
                    <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_SearchBy_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@EntityID", 200, 1,200,EntityID)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,LanguageID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
Set SearchBy = .Execute()
End With
Set Command1 = Nothing
%>
                    <select name="SearchColumn" class="CastlesTextBody">
                      <%
While Not SearchBy.EOF
	If cStr(SearchColumn) = cStr(SearchBy.Fields.Item("SearchByValue").Value) Then
%>
                      <option value="<%=SearchBy.Fields.Item("SearchByValue").Value%>" selected><%=SearchBy.Fields.Item("SearchByName").Value%></option>
                      <%
	Else
%>
                      <option value="<%=SearchBy.Fields.Item("SearchByValue").Value%>"><%=SearchBy.Fields.Item("SearchByName").Value%></option>
                      <%
	End If
	SearchBy.MoveNext()
Wend
SearchBy.Close
Set SearchBy = Nothing
%>
                    </select>
                  </td>
                  <td class="CastlesTextBody" width="150"> 
                    <input type="text" name="SearchValue" class="CastlesTextBody" size="25" maxlength="100" value="<%=SearchValue%>">
                  </td>
                  <td class="CastlesTextBody" width="385" > 
                    <input type="submit" name="Submit" value="<%=WordPhrase_Search%>" class="CastlesTextBlack">
                  </td>
                </tr>
                <tr> 
                  <td colspan="4" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
              </form>
            </table>
          </td>
        </tr>
        <form name="DeleteRecords" method="post" onSubmit="return DeleteCheckedRecords();" action="driveadministrators.asp?DCDataDriverType=SQLMultiDelete&SearchColumn=<%=SearchColumn%>&SearchValue=<%=SearchValue%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>">
          <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
              <tr> 
                <td bgcolor="#000000" colspan="20" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr height="20"> 
                <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                <td bgcolor="<%=TitleBar%>" width="15%" class="CastlesTextHeader"><b><%=WordPhrase_Name%></b></td>
                <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader" align="center"><b><%=WordPhrase_Active%></b></td>
                <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b><%=WordPhrase_DirectLine%></b></td>
                <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b><%=WordPhrase_MobilePhone%></b></td>
                <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                <td bgcolor="<%=TitleBar%>" width="20%" class="CastlesTextHeader"><b><%=WordPhrase_EmailAddress%></b></td>
                <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                <td bgcolor="<%=TitleBar%>" width="15%" class="CastlesTextHeader"><b><%=WordPhrase_Actions%></b></td>
                <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                <td bgcolor="<%=TitleBar%>" width="13%" class="CastlesTextHeader" align="center">&nbsp;&nbsp;
<%
If inStr(EntityModificationAccessString,"(3)" ) <> 0 Then
%>
<b><%=WordPhrase_Delete%></b>
<%
End If
%>
&nbsp;&nbsp;</td>
              </tr>
<%
Response.Flush
if Len(SearchColumn) <> 0 then
	if Not IsNull(SearchResultsReturned) then
		Count = 1
		DeleteCount = 1
		SearchResultsArrayNumRows = uBound(SearchResultsReturned,2)
	
		Field_ID = 0
		Field_AdministratorID = 1
		Field_FirstName = 2
		Field_LastName = 3
		Field_MiddleInitial = 4
		Field_DirectLine = 5
		Field_MobileTelNumber = 6
		Field_EmailAddress = 7
		Field_Active = 8
		Field_MoreRecords = 9

		For SearchResultsArrayRowCounter = 0 to SearchResultsArrayNumRows
			AdministratorID = SearchResultsReturned(Field_AdministratorID,SearchResultsArrayRowCounter)
			FirstName = SearchResultsReturned(Field_FirstName,SearchResultsArrayRowCounter)
			LastName = SearchResultsReturned(Field_LastName,SearchResultsArrayRowCounter)
			MiddleInitial = SearchResultsReturned(Field_MiddleInitial,SearchResultsArrayRowCounter)
			DirectLine = SearchResultsReturned(Field_DirectLine,SearchResultsArrayRowCounter)
			MobileTelNumber = SearchResultsReturned(Field_MobileTelNumber,SearchResultsArrayRowCounter)
			EmailAddress = SearchResultsReturned(Field_EmailAddress,SearchResultsArrayRowCounter)
			Active = SearchResultsReturned(Field_Active,SearchResultsArrayRowCounter)
			MoreRecords = SearchResultsReturned(Field_MoreRecords,SearchResultsArrayRowCounter)
	
			If Count mod 2 = 0 then
				BgColor = "#FFFFFF"
			Else
				BgColor = LightField 
			End If
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td width="1%">&nbsp;</td>
                  <td width="15%" class="CastlesTextBody"><%=LastName%>,&nbsp;<%=FirstName%>&nbsp;<%=MiddleInitial%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody" align="center"> 
                    <%If (Active) = "Y" Then%>
                    <img src="/manage/images/icon_active.gif" width="20" height="20" border="0" alt="Active"> 
                    <%Else%>
                    <img src="/manage/images/icon_nonactive.gif" width="20" height="20" border="0" alt="In-Active"> 
                    <%End If%>
                  </td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody"><%=DirectLine%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody"><%=MobileTelNumber%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="20%" class="CastlesTextBody"><%=EmailAddress%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="15%" class="CastlesTextBody"> 
                    <%
If inStr(EntityModificationAccessString,"(4)") <> 0 Then
%>
                    <a href="viewadministrator.asp?AdministratorID=<%=AdministratorID%>&SearchColumn=<%=Server.URLEncode(SearchColumn)%>&SearchValue=<%=Server.URLEncode(SearchValue)%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>" class="normal"><b><%=WordPhrase_View%></b></a>&nbsp;&nbsp;&nbsp;&nbsp; 
                    <%
End If
If inStr(EntityModificationAccessString,"(2)" ) <> 0 Then
%>
                    <a href="editadministrator.asp?AdministratorID=<%=AdministratorID%>&SearchColumn=<%=Server.URLEncode(SearchColumn)%>&SearchValue=<%=Server.URLEncode(SearchValue)%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>" class="normal"><b><%=WordPhrase_Edit%></b></a>&nbsp;&nbsp;&nbsp;&nbsp; 
                    <%
End If
If inStr(EntityModificationAccessString,"(6)" ) <> 0 Then
FullName = FirstName & " " & LastName
%>
                    <a href="javascript:onClick=PopNotes('1','<%=AdministratorID%>','<%=Server.URLEncode(FullName)%>');" class="normal"><b><%=WordPhrase_Notes%></b></a> 
                    <%
End If
%>
                  </td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="13%" class="CastlesTextBody" align="center">&nbsp;&nbsp;
<%
If inStr(EntityModificationAccessString,"(3)" ) <> 0 Then
%>
                    <input type="checkbox" name="DeleteRecordID<%=DeleteCount%>" value="<%=AdministratorID%>">
<%
End If
%>
                    &nbsp;&nbsp;</td>
                </tr>
                <%
			Count = Count+1
			DeleteCount = DeleteCount+1
		Next
		
		If Count < RecsPerPage Then
			For i = 1 to (((RecsPerPage)+(1))-(Count))
				If Count Mod 2 = 0 then
					Bgcolor = "#FFFFFF"
				Else
					Bgcolor = LightField
				End if
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td colspan="15">&nbsp;</td>
                </tr>
                <%
				Count = Count+1
			Next
		End If
	Else
		For i = 1 to (RecsPerPage)
			If Count Mod 2 = 0 then
				Bgcolor = "#FFFFFF"
			Else
				Bgcolor = LightField
			End if
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td width="1%">&nbsp;</td>
                  <td colspan="15" class="CastlesTextBody"><%=WordPhrase_NoResultsFound%>...</td>
                </tr>
                <%
			Count = Count+1
		Next
	End If
Else
	Count = 1
	For i = 1 to RecsPerPage
		If Count Mod 2 = 0 then
			Bgcolor = "#FFFFFF"
		Else
			Bgcolor = LightField
		End if
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td colspan="15">&nbsp;</td>
                </tr>
                <%
		Count = Count+1
	Next
End If
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <input type="hidden" name="NumberOfRecordsToDelete" value="<%=DeleteCount%>">
              
            </table>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
          <tr> 
            <td class="CastlesTextBody" align="center">&nbsp; 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="50%" class="CastlesTextBody">&nbsp;&nbsp;&nbsp;
                    <%
'Generates Appropriate Pagination Code
If Len(SearchColumn) <> 0 Then
	PageName = "searchadministrators.asp"
	TotalRecords = (MoreRecords)+(RecsPerPage)
	PaginationCode = DCSystemSimpleSearchPagination(Page,WordPhrase_JumpToPage,PageName,MoreRecords,TotalRecords,RecsPerPage,SearchColumn,SearchValue)
	Response.Write PaginationCode
End If
%>
                  </td>
                  <td width="50%" align="right" class="CastlesTextBody">
<%
If inStr(EntityModificationAccessString,"(3)" ) <> 0 Then
%>
                    <input type="submit" name="Submit" value="<%=WordPhrase_Delete%>" class="CastlesTextBlack">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
End If
%>
                  </td>
                </tr>
              </table>
            </td>
        </tr>
		</form>
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
