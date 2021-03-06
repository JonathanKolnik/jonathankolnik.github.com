<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 2
PageTopNavigationSubHeaderID = 8
EntityID = 3
%>
<!--#include virtual="/templates/castlessystemcnekt.asp" -->
<%
'---------------Begin Page Level Multilingual Translation-----------------------
WordPhrasesOnPage = "Actions(|)Active(|)Edit(|)No Results Found(|)Search(|)Search By(|)Search Field(|)System Help Content File(|)System Help Content Name"
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
WordPhrase_Edit = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_NoResultsFound = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Search = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchBy = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SearchField = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SystemHelpContentFile = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_SystemHelpContentName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'---------------End Multilingual Translation-----------------------


'Filters System Help Content
FilterByTopNavigationHeaderID = Request.Form("FilterByTopNavigationHeaderID")
If Len(FilterByTopNavigationHeaderID) = 0 Then
	FilterByTopNavigationHeaderID = Request.QueryString("FilterByTopNavigationHeaderID")
End If

If Len(FilterByTopNavigationHeaderID) <> 0 Then
	Set Command1 = Server.CreateObject("ADODB.Command")
	With Command1	
		.ActiveConnection = Connect
		.CommandText = "Castles_System_SystemHelpContent_Filter_By_Header"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.Parameters.Append .CreateParameter("@TopNavigationHeaderID", 200, 1,200,FilterByTopNavigationHeaderID)
		.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,LanguageID)
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		Set SearchResults = .Execute()
	End With
	Set Command1 = Nothing
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
	TheURL = "/Castles/manage/administration/notes/noteslist.asp?EntityID="+EntityID+"&EntityPrimaryKeyValue="+EntityPrimaryKeyValue+"&EntityName="+EntityName+"&PersonalName="+PersonalName
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
          <td bgcolor="<%=ThinLine%>"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td> 
            <table width="750" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td align="right" colspan="3" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <form name="SearchRecords" method="post" action="searchsystemhelpcontent.asp">
                <tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b><%=FillInSpaceWithNonBreaking(WordPhrase_SearchBy)%>:</b>&nbsp;</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;1&nbsp;]&nbsp;<%=FillInSpaceWithNonBreaking(WordPhrase_SearchField)%>:</td>
                  <td class="CastlesTextBody" width="385">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="65">&nbsp;</td>
                  <td class="CastlesTextBody" width="150"> 
                    <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_SystemHelpContent_SearchBy"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,LanguageID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
Set SearchBy = .Execute()
End With
Set Command1 = Nothing
%>
                    <select name="FilterByTopNavigationHeaderID" class="CastlesTextBody">
                      <%
While Not SearchBy.EOF
	If cStr(FilterByTopNavigationHeaderID) = cStr(SearchBy.Fields.Item("TopNavigationHeaderID").Value) Then
%>
                      <option value="<%=SearchBy.Fields.Item("TopNavigationHeaderID").Value%>" selected><%=SearchBy.Fields.Item("TopNavigationHeaderName").Value%> -(For&nbsp;<%=SearchBy.Fields.Item("SystemLoginTypeName").Value%>)</option>
                      <%
	Else
%>
                      <option value="<%=SearchBy.Fields.Item("TopNavigationHeaderID").Value%>"><%=SearchBy.Fields.Item("TopNavigationHeaderName").Value%> - (For&nbsp;<%=SearchBy.Fields.Item("SystemLoginTypeName").Value%>)</option>
                      <%
	End If
	SearchBy.MoveNext()
Wend
SearchBy.Close
Set SearchBy = Nothing
%>
                    </select>
                  </td>
                  <td class="CastlesTextBody" width="385" > 
                    <input type="submit" name="Submit" value="<%=WordPhrase_Search%>" class="CastlesTextBlack">
                  </td>
                </tr>
                <tr> 
                  <td colspan="3" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
              </form>
            </table>
          </td>
        </tr>
          <tr> 
            <td> 
              
            <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
              <tr> 
                  <td bgcolor="#000000" colspan="20" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr height="20"> 
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  
                <td bgcolor="<%=TitleBar%>" width="20%" class="CastlesTextHeader"><b><%=WordPhrase_SystemHelpContentName%></b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  
                <td bgcolor="<%=TitleBar%>" width="20%" class="CastlesTextHeader" align="left"><b><%=WordPhrase_SystemHelpContentFile%></b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b><%=WordPhrase_Actions%></b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  
                <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  
                <td bgcolor="<%=TitleBar%>" width="5%" class="CastlesTextHeader">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="13%" class="CastlesTextHeader" align="center">&nbsp;&nbsp;</td>
                </tr>
                <%
Response.Flush
If Len(FilterByTopNavigationHeaderID) <> 0 Then
	If Not SearchResults.EOF Then
		Count = 1
		DeleteCount = 1
		SearchResultsArray = SearchResults.getrows
		SearchResults.close
		Set SearchResults = Nothing
		SearchResultsArrayNumRows = uBound(SearchResultsArray,2)
	
		Field_SystemHelpContentID = 0
		Field_SystemHelpContentName = 1
		Field_SystemHelpContentFileName = 2

		For SearchResultsArrayRowCounter = 0 to SearchResultsArrayNumRows
			SystemHelpContentID = SearchResultsArray(Field_SystemHelpContentID,SearchResultsArrayRowCounter)
			SystemHelpContentName = SearchResultsArray(Field_SystemHelpContentName,SearchResultsArrayRowCounter)
			SystemHelpContentFileName = SearchResultsArray(Field_SystemHelpContentFileName,SearchResultsArrayRowCounter)
	
			If Count mod 2 = 0 then
				BgColor = "#FFFFFF"
			Else
				BgColor = LightField 
			End If
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td width="1%">&nbsp;</td>
                  
                <td width="20%" class="CastlesTextBody"><%=SystemHelpContentName%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  
                <td width="20%" class="CastlesTextBody" align="left"><%=SystemHelpContentFileName%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody"><a href="editsystemhelpcontent.asp?SystemHelpContentID=<%=SystemHelpContentID%>&FilterByTopNavigationHeaderID=<%=FilterByTopNavigationHeaderID%>&SystemHelpContentLanguageID=<%=LanguageID%>&SystemHelpContentName=<%=Server.URLEncode(SystemHelpContentName)%>" class="normal"><b><%=WordPhrase_Edit%></b></a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
				  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody">&nbsp;</td>
                  <td width="1%" align="center">&nbsp;</td>
                  
                <td width="10%" class="CastlesTextBody">&nbsp;</td>
                  <td width="1%" align="center">&nbsp;</td>
                  
                <td width="5%" class="CastlesTextBody">&nbsp; </td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="13%" class="CastlesTextBody" align="center">&nbsp;&nbsp;&nbsp;</td>
                </tr>
                <%
			Count = Count+1
			DeleteCount = DeleteCount+1
		Next
		
		If Count < 15 Then
			For i = 1 to (((15)+(1))-(Count))
				If Count Mod 2 = 0 then
					Bgcolor = "#FFFFFF"
				Else
					Bgcolor = LightField
				End if
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td colspan="15">&nbsp;</td>
                </tr>
                <%
				Count = Count+1
			Next
		End If
	Else
		For i = 1 to (15)
			If Count Mod 2 = 0 then
				Bgcolor = "#FFFFFF"
			Else
				Bgcolor = LightField
			End if
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
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
	For i = 1 to 15
		If Count Mod 2 = 0 then
			Bgcolor = "#FFFFFF"
		Else
			Bgcolor = LightField
		End if
%>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
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
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="15" height="1"><img src="/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
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
                  </td>
                  <td width="50%" align="right" class="CastlesTextBody">&nbsp; </td>
                </tr>
              </table>
            </td>
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
