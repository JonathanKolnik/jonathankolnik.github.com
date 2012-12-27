<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationSubHeaderID = Request.QueryString("PageTopNavigationSubHeaderID")
%>
<!--#include virtual="/templates/Castlessystemcnektonly.asp" -->
<%

'---------------Begin Page Level Multilingual Translation-----------------------
WordPhrasesOnPage = "Actions(|)Create New Note(|)Created By(|)Created Date(|)Delete(|)Hide Help Text(|)Jump To Page(|)No Results Found(|)Note Title(|)Notes For(|)Search(|)Search By(|)Search Field(|)Search Value(|)Show Help Text(|)View"
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
WordPhrase_CreateNewNote  = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_CreatedBy  = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_CreatedDate = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Delete = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_HideHelpText = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_JumpToPage = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_NoResultsFound = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_NoteTitle = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_NotesFor = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Search = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchBy = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchField = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchValue = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_ShowHelpText = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_View = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'---------------End Multilingual Translation-----------------------

RecsPerPage = 10

Page = Request.QueryString("Page")
TotalRecords = Request.QueryString("TotalRecords")
MoreRecords = Request.QueryString("MoreRecords")
LastSearchValue = Request.Form("LastSearchValue")

EntityID = Request.QueryString("EntityID")
EntityPrimaryKeyValue = Request.QueryString("EntityPrimaryKeyValue")
EntityName = Request.QueryString("EntityName")
PersonalName = Request.QueryString("PersonalName")

If Len(Request.Form("SearchColumn")) <> 0 Then
	SearchColumn = Request.Form("SearchColumn")
	SearchValue = Request.Form("SearchValue")
Else
	SearchColumn = Request.QueryString("SearchColumn")
	SearchValue = Request.QueryString("SearchValue")
End If

If Len(Page) = 0 Then
	Page = 1
End If

If Len(RecsPerPage) = 0 Then
	RecsPerPage = 10
End If

If Len(SearchColumn) = 0 Then
	SearchColumn = "NoteTitle"
	SearchValue = "%"
End If

Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_Notes_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@Page", 200, 1,200,Page)
	.Parameters.Append .CreateParameter("@RecsPerPage", 200, 1,200,RecsPerPage)
	.Parameters.Append .CreateParameter("@EntityID", 200, 1,200,EntityID)
	.Parameters.Append .CreateParameter("@EntityPrimaryKeyValue", 200, 1,200,EntityPrimaryKeyValue)
	.Parameters.Append .CreateParameter("@SearchColumn", 200, 1,200,SearchColumn)
	.Parameters.Append .CreateParameter("@SearchValue", 200, 1,200,SearchValue)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set SearchResults = .Execute()
End With
Set Command1 = Nothing

If SearchValue = "%" Then
	SearchValue = Replace(SearchValue,"%","")
End If
%>
<html><!-- #BeginTemplate "/Templates/CastlesSystemPopUp.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Castles - Management System</title>
<!-- #EndEditable -->
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
function DeleteCheckedRecords(){
	if(confirm("Are you sure you want to permanently delete the selected item(s)?")){
		return true;
	}else {
		return false;
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
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="190"><img src="/manage/images/castles_logo.GIF" width="190" height="40" usemap="#Map" border="0"></td>
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
          <td bgcolor="<%=NavDarkBorder%>" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavRegular%>" height="19"><img src="/manage/images/clear10pixel.gif" width="1" height="19"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavDarkBorder%>" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
<!-- #BeginEditable "topnav" --> 
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="19"> 
            <table width="300" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="CastlesTextNavDark"> 
                  <%
If Request.Cookies("CreateNotes") = "Y" Then
%>
                  &nbsp;&nbsp;&nbsp;&gt;&nbsp;<a href="createnote.asp?EntityID=<%=EntityID%>&EntityPrimaryKeyValue=<%=EntityPrimaryKeyValue%>&EntityName=<%=Server.URLEncode(EntityName)%>&PersonalName=<%=Server.URLEncode(PersonalName)%>&PageTopNavigationSubHeaderID=<%=PageTopNavigationSubHeaderID%>" class="navdark">Create 
                  New Note</a> 
                  <%
End If
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <!-- #EndEditable -->
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
<!-- #BeginEditable "body" --> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
            <table width="700" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td class="CastlesTextMid" colspan="3" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td class="CastlesTextMid" width="3">&nbsp;</td>
                <td class="CastlesTextBodyBig" width="600"><b><%=UCase(WordPhrase_NotesFor)%>&nbsp;-&nbsp;<%=UCase(EntityName)%>&nbsp; 
                  <%
If Len(PersonalName) <> 0 Then
%>
                  -&nbsp;<%=UCase(PersonalName)%>&nbsp; 
                  <%
End If
%>
                  -&nbsp;#<%=EntityPrimaryKeyValue%>&nbsp;&nbsp;&nbsp;</b></td>
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
                <td class="CastlesTextBody" width="100" align="right"><a href="<%=SelfLink%>" class="normal"><b><%=WordPhrase_ShowHelpText%></b></a></td>
                <%
Else
	SelfLink = SelfLink & "HideHelpContent=Y"
%>
                <td class="CastlesTextBody" width="100" align="right"><a href="<%=SelfLink%>" class="normal"><b><%=WordPhrase_HideHelpText%></b></a></td>
                <%
End If
%>
              </tr>
              <%
If SystemHelpContentText <> "" Then
%>
              <tr> 
                <td class="CastlesTextMid" width="3">&nbsp;</td>
                <td class="CastlesTextBody" colspan="2"><%=SystemHelpContentText%></td>
              </tr>
              <%
End If
%>
              <tr> 
                <td class="CastlesTextMid" colspan="3" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="<%=ThinLine%>"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td> 
            <table width="600" border="0" cellspacing="0" cellpadding="3">
              <form name="SearchNotes" method="post" action="noteslist.asp?EntityID=<%=EntityID%>&EntityPrimaryKeyValue=<%=EntityPrimaryKeyValue%>&EntityName=<%=Server.URLEncode(EntityName)%>&PersonalName=<%=Server.URLEncode(PersonalName)%>">
                <tr> 
                  <td colspan="4" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b><%=FillInSpaceWithNonBreaking(WordPhrase_SearchBy)%>:</b>&nbsp;</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;1&nbsp;]&nbsp;<%=FillInSpaceWithNonBreaking(WordPhrase_SearchField)%>:</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;2&nbsp;]&nbsp;<%=FillInSpaceWithNonBreaking(WordPhrase_SearchValue)%>:</td>
                  <td class="CastlesTextBody" width="235">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="65">&nbsp;</td>
                  <td class="CastlesTextBody" width="150"> 
                    <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_Notes_SearchBy"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
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
                  <td class="CastlesTextBody" width="235" > 
                    <input type="submit" name="Submit" value="Search" class="CastlesTextBlack">
                  </td>
                </tr>
                <tr> 
                  <td colspan="4" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
              </form>
            </table>
          </td>
        </tr>
        <form name="DeleteNotes" onSubmit="return DeleteCheckedRecords();" method="post" action="DriveNotes.asp?castkesdcdatadriverType=SQLMultiDelete&SearchColumn=<%=Server.URLEncode(SearchColumn)%>&SearchValue=<%=Server.URLEncode(SearchValue)%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>&EntityID=<%=EntityID%>&EntityPrimaryKeyValue=<%=EntityPrimaryKeyValue%>&EntityName=<%=Server.URLEncode(EntityName)%>&PersonalName=<%=Server.URLEncode(PersonalName)%>">
          <tr> 
            <td> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#000000"> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="20%" class="CastlesTextHeader"><b>Created 
                    Date </b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="25%" class="CastlesTextHeader"><b>Created 
                    By </b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="30%" class="CastlesTextHeader"><b>Note 
                    Title</b> </td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b>Actions</b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader" align="center">&nbsp;&nbsp; 
                    <%
If Request.Cookies("DeleteNotes") = "Y" Then
%>
                    <b>Delete</b> 
                    <%
End If
%>
                    &nbsp;&nbsp;</td>
                </tr>
                <%
Response.Flush
if Not SearchResults.EOF then
	Count = 1
	DeleteCount = 1
	SearchResultsArray = SearchResults.getrows
	SearchResults.close
	Set SearchResults = Nothing
	SearchResultsArrayNumRows = uBound(SearchResultsArray,2)

	Field_ID = 0
	Field_NoteID = 1
	Field_NoteTitle = 2
	Field_CreatedBySystemLoginID = 3
	Field_CreatedDateTime = 4
	Field_MoreRecords = 5

	For SearchResultsArrayRowCounter = 0 to SearchResultsArrayNumRows
		NoteID = SearchResultsArray(Field_NoteID,SearchResultsArrayRowCounter)
		NoteTitle = SearchResultsArray(Field_NoteTitle,SearchResultsArrayRowCounter)
		CreatedBySystemLoginID = SearchResultsArray(Field_CreatedBySystemLoginID,SearchResultsArrayRowCounter)
		CreatedDateTime = SearchResultsArray(Field_CreatedDateTime,SearchResultsArrayRowCounter)
		MoreRecords = SearchResultsArray(Field_MoreRecords,SearchResultsArrayRowCounter)

		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_SystemLogin_Name_Type"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@SystemLoginID", 200, 1,200,CreatedBySystemLoginID)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set CreatedBySystemLoginInfo = .Execute()
		End With
		Set Command1 = Nothing
		CreatedByInfo = CreatedBySystemLoginInfo.Fields.Item("FirstName").Value & "&nbsp;" & CreatedBySystemLoginInfo.Fields.Item("LastName").Value & "&nbsp;(" & CreatedBySystemLoginInfo.FIelds.Item("SystemLoginTypeName").Value & ")"

		If Count mod 2 = 0 then
			BgColor = "#FFFFFF"
		Else
			BgColor = LightField 
		End If
%>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td width="1%">&nbsp;</td>
                  <td width="20%" class="CastlesTextBody"><%=CreatedDateTime%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="25%" class="CastlesTextBody"><%=CreatedByInfo%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="30%" class="CastlesTextBody"><%=NoteTitle%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody"> 
                    <a href="viewnote.asp?NoteID=<%=NoteID%>&SearchColumn=<%=Server.URLEncode(SearchColumn)%>&SearchValue=<%=Server.URLEncode(SearchValue)%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>&EntityID=<%=EntityID%>&EntityPrimaryKeyValue=<%=EntityPrimaryKeyValue%>&EntityName=<%=Server.URLEncode(EntityName)%>&PersonalName=<%=Server.URLEncode(PersonalName)%>&PageTopNavigationSubHeaderID=<%=PageTopNavigationSubHeaderID%>" class="normal"><b><%=WordPhrase_View%></b></a> 
                  </td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody" align="center">&nbsp;&nbsp; 
                    <%
If Request.Cookies("DeleteNotes") = "Y" Then
%>
                    <input type="checkbox" name="DeleteRecordID<%=DeleteCount%>" value="<%=NoteID%>">
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
			If Count mod 2 = 0 then
				BgColor = "#FFFFFF"
			Else
				BgColor = LightField 
			End If
%>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td colspan="10">&nbsp;</td>
                </tr>
                <%
			Count = Count+1
		Next
	End If
Else
Count = 1
	For i = 1 to (RecsPerPage)
		If Count mod 2 = 0 then
			BgColor = "#FFFFFF"
		Else
			BgColor = LightField 
		End If
%>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td width="1%">&nbsp;</td>
                  <td colspan="11" class="CastlesTextBody"><%=WordPhrase_NoResultsFound%>...</td>
                </tr>
                <%
		Count = Count+1
	Next
End If
%>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="11" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <input type="hidden" name="NumberOfRecordsToDelete" value="<%=DeleteCount%>">
              </table>
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td class="CastlesTextBody" align="left"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="50%" class="CastlesTextBody">&nbsp;&nbsp;&nbsp; 
                    <%
'Generates Appropriate Pagination Code
If Len(SearchColumn) <> 0 Then
	PageName = "noteslist.asp"
	TotalRecords = (MoreRecords)+(RecsPerPage)
	PaginationCode = DCSystemSimpleSearchPagination(Page,WordPhrase_JumpToPage,PageName,MoreRecords,TotalRecords,RecsPerPage,SearchColumn,SearchValue,EntityID,EntityPrimaryKeyValue,EntityName,PersonalName)
	Response.Write PaginationCode
End If
Function DCSystemSimpleSearchPagination(Page,WordPhrase_JumpToPage,PageName,MoreRecords,TotalRecords,RecsPerPage,SearchColumn,SearchValue,EntityID,EntityPrimaryKeyValue,EntityName,PersonalName)
	If Page <> 1 Then
		TotalRecords = Request.QueryString("TotalRecords")
	End If

	Paginations = (TotalRecords)/(RecsPerPage)
	If TotalRecords <= RecsPerPage Then
		Paginations = 1
	Else
		If (TotalRecords) Mod (RecsPerPage) <> 0 Then
			LeftOver = (TotalRecords)mod(RecsPerPage)
			Paginations = cStr(Paginations)
			DecimalLocation = inStr(Paginations,".")
			Paginations = Mid(Paginations,1,DecimalLocation-1)
			Paginations = cLng(Paginations)
			Paginations = (Paginations)+(1)
		End If
	End If

	PageCount = 1
	DCSystemSimpleSearchPagination = WordPhrase_JumpToPage & ": "

	For i = 1 to Paginations 
		If cStr(PageCount) <> cStr(Page) Then
			DCSystemSimpleSearchPagination = DCSystemSimpleSearchPagination &  "<a href=" & DQ & PageName & "?SearchColumn=" & Server.URLEncode(SearchColumn) & "&SearchValue=" & Server.URLEncode(SearchValue) & "&Page=" & PageCount & "&TotalRecords=" & TotalRecords & "&EntityID=" & EntityID & "&EntityPrimaryKeyValue=" & EntityPrimaryKeyValue & "&EntityName=" & Server.URLEncode(EntityName) & "&PersonalName=" & Server.URLEncode(PersonalName) &DQ &" class=""normal""><b>" & PageCount & "</b></a>|"
		Else
			DCSystemSimpleSearchPagination = DCSystemSimpleSearchPagination & "<b>" & PageCount & "</b>|" 
		End If
		PageCount = PageCount + 1
	Next
End Function
%>
                  </td>
                  <td width="50%" align="right" class="CastlesTextBody"> 
                    <%
If Request.Cookies("DeleteNotes") = "Y" Then
%>
                    <input type="submit" name="Submit" value="<%=WordPhrase_Delete%>" class="CastlesTextBlack">
                    &nbsp;&nbsp;&nbsp; 
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
      <!-- #EndEditable -->
      </td>
  </tr>
  <tr width="100%"> 
    <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
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
</body>
<!-- #EndTemplate --></html>
