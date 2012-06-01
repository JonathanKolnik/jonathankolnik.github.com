<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 6
PageTopNavigationSubHeaderID = 22
EntityID = 10
%>
<!--#include virtual="/templates/castlessystemcnekt.asp" -->
<!--#include virtual="/templates/castlesdcsystemsimplesearch.asp" -->
<%
'---------------Begin Page Level Multilingual Translation-----------------------
WordPhrasesOnPage = "Actions(|)Active(|)Company Name(|)Delete(|)Designer(|)Direct Line(|)Edit(|)Email Address(|)Jump To Page(|)Mobile Phone(|)Name(|)No Results Found(|)Notes(|)Search(|)Search By(|)Search Field(|)Search Value(|)View"
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
	.Parameters.Append .CreateParameter("@WhereClause",201,1,20000,WhereClause)
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
WordPhrase_CompanyName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Delete = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Broker = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_TelNumber = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
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

'Page = Request.QueryString("Page")
'TotalRecords = Request.QueryString("TotalRecords")
'MoreRecords = Request.QueryString("MoreRecords")
'LastSearchValue = Request.Form("LastSearchValue")
'SearchOnEntity = "Listings"

' for pagination stuff at bottom
RecsPerPage = 10

Page = Request.QueryString("Page")
If Len(Page) = 0 Then
	Page = 1
End If

TotalRecords = Request.QueryString("TotalRecords")
MoreRecords = Request.QueryString("MoreRecords")
' end pagination stuff

Function findValue(src)
	dest = request.Form(src)
	if Len(dest)=0 then
		dest = request.QueryString(src)
	end if
	findValue = dest
End Function

SearchListingID = findValue("SearchListingID")
SearchedStates = findValue("SearchedStates")
'response.write SearchedAreas
SearchedSizes = findValue("SearchedSizes")
SearchedAddress = findValue("SearchedAddress")
'SearchedApartmentStatus = findValue("SearchedApartmentStatus")
PriceFrom = findValue("PriceFrom")
PriceTo = findValue("PriceTo")

if inStr(SearchedStates,"ALL") <> 0 then
	StatesNoFilter = true
else
	StatesNoFilter = false
	StatesReplaced = Replace(SearchedStates,"(","")
	StatesReplaced = Replace(StatesReplaced,")","")
end if

if inStr(SearchedSizes,"ALL") <> 0 then
	SizesNoFilter = true
else
	SizesNoFilter = false
	SizesReplaced = Replace(SearchedSizes,"(","")
	SizesReplaced = Replace(SizesReplaced,")","")
end if
	
if SearchedApartmentStatus = "ALL" then
	StatusNoFilter = true
end if

if (SearchListingID<>"")then
	set Command1 = Server.CreateObject("ADODB.Command")
		with command1	
			.ActiveConnection = connect
			.CommandText = "Castles_System_Listings_SearchbyID_ForDesigner"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@ListingID", 200, 1,200,SearchListingID)
			.Parameters.Append .CreateParameter("@Page", 200, 1,200,Page)
			.Parameters.Append .CreateParameter("@RecsPerPage", 200, 1,200,RecsPerPage)	
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			set Listings = .Execute()
		end with
	set command1 = nothing
	if((not Listings.BOF)and(not Listings.EOF))then
		ListingsHasRecords = true
	else
		ListingsHasRecords = false
	end if
end if
'response.write ClientID
if ((SearchListingID="")AND(SearchedAddress<>""))then
	set Command1 = Server.CreateObject("ADODB.Command")
		with command1	
			.ActiveConnection = connect
			.CommandText = "Castles_System_Listings_SearchbyAddress_ForDesigner"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@Page", 200, 1,200,Page)
			.Parameters.Append .CreateParameter("@RecsPerPage", 200, 1,200,RecsPerPage)
			.Parameters.Append .CreateParameter("@Address", 200, 1,200,SearchedAddress)	
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			set Listings = .Execute()
		end with
	set command1 = nothing
	if((not Listings.BOF)and(not Listings.EOF))then
		ListingsHasRecords = true
	else
		ListingsHasRecords = false
	end if
end if
if ((SearchListingID="")and(SearchedAddress="")and(SearchedStates<>"")and(SearchedSizes<>"")) then
	if NOT StatesNoFilter then
		StatesArray = Split(StatesReplaced,",")
	end if
	if NOT SizesNoFilter then	
		SizesArray = Split(SizesReplaced,",")
	end if
	counter = 1

	if (NOT StatusNoFilter AND NOT SizesNoFilter AND NOT StatesNoFilter) then	
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.StateProvinceID ="	
		for each z in SizesArray 
			for each x in StatesArray
				if counter = 1 then
					SQLStmt = SQLStmt & Trim(x) 
				else
					SQLStmt = SQLStmt & " OR Castles_Listings.Deleted <> 'Y' AND Castles_Listings.StateProvinceID = " & Trim(x) & " AND Castles_Listings.ListingPublishStatusID = 4"
				end if				
				SQLStmt = SQLStmt & " AND Castles_Listings.PropertyTypeID = " & Trim(z) & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
				counter = counter + 1
			Next
		Next
	end if
	
	if (NOT StatusNoFilter AND NOT SizesNoFilter AND StatesNoFilter) then	
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.PropertyTypeID ="	
		for each z in SizesArray 
			if counter = 1 then
				SQLStmt = SQLStmt & Trim(z) 
			else
				SQLStmt = SQLStmt & " OR Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4"	
			end if				
			SQLStmt = SQLStmt & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
			counter = counter + 1
		Next
	end if
	
	if (StatusNoFilter AND NOT SizesNoFilter AND StatesNoFilter) then	
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.PropertyTypeID ="	
		for each z in SizesArray 
			if counter = 1 then
				SQLStmt = SQLStmt &  Trim(z)
			else
				SQLStmt = SQLStmt & " OR Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.PropertyTypeID = " & Trim(z) 	
			end if				
			SQLStmt = SQLStmt & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
			counter = counter + 1
		Next
	end if
	
	if (NOT StatusNoFilter AND SizesNoFilter AND StatesNoFilter) then	
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4"					
		SQLStmt = SQLStmt & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
	end if
	
	if (StatusNoFilter AND SizesNoFilter AND StatesNoFilter) then	
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4"				
		SQLStmt = SQLStmt & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
	end if
	
	if (NOT StatusNoFilter AND SizesNoFilter AND NOT StatesNoFilter) then	
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.StateProvinceID = "	
		for each x in StatesArray
			if counter = 1 then
				SQLStmt = SQLStmt & Trim(x)
			else
				SQLStmt = SQLStmt & " OR Castles_Listings.Deleted <> 'Y' AND Castles_Listings.StateProvinceID = " & Trim(x) & " AND Castles_Listings.ListingPublishStatusID = 4"
			end if				
			SQLStmt = SQLStmt & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
			counter = counter + 1
		Next
	end if
	
	if (StatusNoFilter AND NOT SizesNoFilter AND NOT StatesNoFilter) then
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.StateProvinceID = "	
		for each z in SizesArray 
			for each x in StatesArray
				if counter = 1 then
					SQLStmt = SQLStmt & Trim(x) 
				else
					SQLStmt = SQLStmt & " OR Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.StateProvinceID = " & Trim(x)	
				end if				
				SQLStmt = SQLStmt & " AND Castles_Listings.PropertyTypeID = " & Trim(z) & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
				counter = counter + 1
			Next
		Next
	end if
	
	if (StatusNoFilter AND SizesNoFilter AND NOT StatesNoFilter) then
		SQLStmt = SQLStmt & " Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.StateProvinceID = "	 
		for each x in StatesArray
			if counter = 1 then
				SQLStmt = SQLStmt & Trim(x) 
			else
				SQLStmt = SQLStmt & " OR Castles_Listings.Deleted <> 'Y' AND Castles_Listings.ListingPublishStatusID = 4 AND Castles_Listings.StateProvinceID = " & Trim(x) 	
			end if				
			SQLStmt = SQLStmt & " AND Castles_Listings.ListPrice >= " & PriceFrom & " AND Castles_Listings.ListPrice <= " & PriceTo
			counter = counter + 1
		Next
	end if
	
	SQLStmt = SQLStmt & " ORDER BY Castles_Listings.ListPrice"
	
	'response.Write SQLStmt
	'Response.Write "<br><br><br>StatesArray = " & StatesReplaced & "<br><br><br>"
	
	set Command1 = Server.CreateObject("ADODB.Command")
		with command1	
			.ActiveConnection = connect
			.CommandText = "Castles_System_Listings_Manage_Search_ForDesigner"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@Page", 200, 1,200,Page)
			.Parameters.Append .CreateParameter("@RecsPerPage", 200, 1,200,RecsPerPage)
			.Parameters.Append .CreateParameter("@Whereclause", 201, 1,20000,SQLStmt)	
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			set Listings = .Execute()
		end with
	set command1 = nothing
	if((not Listings.BOF)and(not Listings.EOF))then
		ListingsHasRecords = true
	else
		ListingsHasRecords = false
	end if
	'response.write "ListingsHasRecords: " & ListingsHasRecords
end if

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
	EntityName = "<%=WordPhrase_Broker%>"
	TheURL = "/manage/administration/notes/noteslist.asp?EntityID="+EntityID+"&EntityPrimaryKeyValue="+EntityPrimaryKeyValue+"&EntityName="+EntityName+"&PersonalName="+PersonalName
	WinName = "NotesList"+EntityPrimaryKeyValue
	Features = "width=720,height=608,resizable,scrollbars=yes"
	window.open(TheURL,WinName,Features);
}  


function PickUpCheckedRecords(){
	if(confirm("Are you sure you want to change the status of the selected item(s)?")){
		return true;
	}else {
		return false;
	}
}

function Search(){
	var ErrorString = ""
	var ErrorTrue = ""
	
	//if (document.SearchRecords.SearchValue.value == "") {
		//ErrorString = ErrorString + "Please enter a value for the search. \r"
		//ErrorTrue="Y"
	//}
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
                <td width="190"><img src="../../images/castles_logo.GIF" width="190" height="40" usemap="#Map" border="0"></td>
                <td width="560"> 
                  <table width="560" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td class="CastlesTextBody" width="230" align="right"><%=FillInSpaceWithNonBreaking(WordPhrase_Welcome)%>,&nbsp;<%=SystemLoginNickName%>&nbsp;&nbsp;<b>&gt;&nbsp;<a href="../../editProfile.asp" class="normal"><%=FillInSpaceWithNonBreaking(WordPhrase_EditProfile)%></a></b>&nbsp;&nbsp;<b>&gt;&nbsp;<a href="../../logout.asp" class="normal"><%=FillInSpaceWithNonBreaking(WordPhrase_LogOut)%></a>&nbsp;&nbsp;&nbsp;</b></td>
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
<%
set Command1 = Server.CreateObject("ADODB.Command")
with command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_StateProvinceAndCountry_ABV_Admin_DynamicList"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	set StateCountry = .Execute()
end with
set Command1 = nothing
set Command1 = Server.CreateObject("ADODB.Command")
	with command1	
		.ActiveConnection = connect
		.CommandText = "Castles_System_PropertyTypes_List"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		set Sizes = .Execute()
	end with
set Command1 = nothing
set Command1 = Server.CreateObject("ADODB.Command")
	with command1	
		.ActiveConnection = connect
		.CommandText = "Castles_System_ListingStatus_List"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		set ListingStatus = .Execute()
	end with
set Command1 = nothing
set Command1 = Server.CreateObject("ADODB.Command")
	with command1	
		.ActiveConnection = connect
		.CommandText = "Castles_System_Listings_Prices_List"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		set Prices = .Execute()
	end with
set Command1 = nothing
%>		   
            <table width="750" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td align="right" colspan="4" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <form name="SearchRecords" method="post" onSubmit="return Search();" action="searchadvertisements.asp">
                <tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b><%=FillInSpaceWithNonBreaking(WordPhrase_SearchBy)%>:</b>&nbsp;</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;1&nbsp;]&nbsp;Listing ID:</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;2&nbsp;]&nbsp;Address:</td>
                  <td class="CastlesTextBody" width="385">&nbsp;</td>
                </tr>
				<tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b></b>&nbsp;</td>
                  <td class="CastlesTextBody" width="150"> 
                    <input type="text" name="SearchListingID" class="CastlesTextBody" value="<%=SearchListingID%>">
                  </td>
                  <td class="CastlesTextBody" width="150"> 
                    <input type="text" name="SearchedAddress" class="CastlesTextBody" value="<%=SearchedAddress%>">
                  </td>
                  <td class="CastlesTextBody" width="385">&nbsp; </td>
                </tr>
				<tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b></b>&nbsp;</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;4&nbsp;]&nbsp;Price 
                    From:</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;5&nbsp;]&nbsp;Price 
                    To:</td>
                  <td class="CastlesTextBody" width="385">&nbsp;</td>
                </tr>
				<tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b></b></td>
                  <td class="CastlesTextBody" width="150"> 
                    <select name="PriceFrom" class="CastlesTextBody">
                      <option value="0">Any Price</option>
                      <%
While (NOT Prices.EOF)
	if Trim(Prices.Fields.Item("OptionValue").Value) = PriceFrom then%>
                      <option value="<%=Trim(Prices.Fields.Item("OptionValue").Value)%>" selected><%=(Prices.Fields.Item("OptionName").Value)%></option>
                      <%else%>
                      <option value="<%=Trim(Prices.Fields.Item("OptionValue").Value)%>"><%=(Prices.Fields.Item("OptionName").Value)%></option>
                      <%end if
Prices.MoveNext()
Wend
Prices.close
%>
                    </select>
                  </td>
                  <td class="CastlesTextBody" width="150"> 
                    <select name="PriceTo" class="CastlesTextBody">
                      <option value="100000000">Any Price</option>
                      <%
set Command1 = Server.CreateObject("ADODB.Command")
	with command1	
		.ActiveConnection = connect
		.CommandText = "Castles_System_Listings_Prices_List"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		set Prices = .Execute()
	end with
set Command1 = nothing

While (NOT Prices.EOF)
	if Trim(Prices.Fields.Item("OptionValue").Value) = PriceTo then%>
                      <option value="<%=Trim(Prices.Fields.Item("OptionValue").Value)%>" selected><%=(Prices.Fields.Item("OptionName").Value)%></option>
                      <%else%>
                      <option value="<%=Trim(Prices.Fields.Item("OptionValue").Value)%>"><%=(Prices.Fields.Item("OptionName").Value)%></option>
                      <%end if
Prices.MoveNext()
Wend
Prices.close
%>
                    </select>
                  </td>
                  <td class="CastlesTextBody" width="385">&nbsp;</td>
                </tr>
				<tr> 
                  <td align="right" width="65" class="CastlesTextBody"><b></b>&nbsp;</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;6&nbsp;]&nbsp;Sizes:</td>
                  <td class="CastlesTextBody" width="150">[&nbsp;7&nbsp;]&nbsp;States:</td>
                  <td class="CastlesTextBody" width="385">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="65">&nbsp;</td>
                  <td class="CastlesTextBody" width="150"> 
                    <select name="SearchedSizes" class="CastlesTextBody" multiple>
                      <%
if ((inStr(SearchedSizes,"ALL") <> 0) OR (SearchedSizes = "")) then
%>
                      <option value="ALL" selected>All Types</option>
                      <%
else
%>
                      <option value="ALL">All Types</option>
                      <%
end if
%>
                      <%
While (NOT Sizes.EOF)
	SizesString = "(" & cStr(Sizes.Fields.Item("PropertyTypeID").Value) & ")"
	if inStr(SearchedSizes,SizesString) <> 0 then %>
                      <option value="(<%=(Sizes.Fields.Item("PropertyTypeID").Value)%>)" selected><%=(Sizes.Fields.Item("PropertyTypeName").Value)%></option>
                      <%else%>
                      <option value="(<%=(Sizes.Fields.Item("PropertyTypeID").Value)%>)"><%=(Sizes.Fields.Item("PropertyTypeName").Value)%></option>
                      <%end if
Sizes.MoveNext()
Wend
Sizes.close
%>
                    </select>
                  </td>
                  <td class="CastlesTextBody" width="150"> 
                    <select name="SearchedStates" class="CastlesTextBody" multiple>
                      <%
if ((inStr(SearchedStates,"ALL") <> 0) OR (SearchedStates = "")) then
%>
                      <option value="ALL" selected>All States</option>
                      <%
else
%>
                      <option value="ALL">All States</option>
                      <%
end if

While Not StateCountry.EOF
	StateString = "(" & cStr(StateCountry.Fields.Item("StateProvinceID").Value) & ")"
	if inStr(SearchedStates,StateString) <> 0 then
		makeSelected = "selected"
	else
		makeSelected = ""
	end if
%>
                      <option value="(<%=StateCountry.Fields.Item("StateProvinceID").Value%>)" <%=makeSelected%>><%=StateCountry.Fields.Item("StateProvinceAbv").Value%> / <%=StateCountry.Fields.Item("CountryAbv").Value%></option>
                      <%
	StateCountry.MoveNext()
Wend
StateCountry.close
%>
                    </select>
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
        <form name="PickUpRecords" method="post" onSubmit="return PickUpCheckedRecords();" action="driveAdvertisementsPickUp.asp?SearchListingID=<%=SearchListingID%>&SearchedAddress=<%=SearchedAddress%>&SearchedSizes=<%=SearchedSizes%>&SearchedStates=<%=SearchedStates%>&SearchedStatus=<%=SearchedApartmentStatus%>&PriceFrom=<%=PriceFrom%>&PriceTo=<%=PriceTo%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>">
          <tr> 
            <td> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
                <tr> 
                  <td bgcolor="#000000" colspan="20" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr height="20"> 
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b>Broker/Company</b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b>Address</b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="5%" class="CastlesTextHeader"><b>Available</b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b>Phone</b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="15%" class="CastlesTextHeader"><b>Email</b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader"><b>Actions</b></td>
                  <td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
                  <td bgcolor="<%=TitleBar%>" width="11%" class="CastlesTextHeader" align="center">&nbsp;&nbsp; 
                    <%
If inStr(EntityModificationAccessString,"(8)" ) <> 0 Then
%>
                    <b>Pick-Up</b> 
                    <%
End If
%>
                    &nbsp;&nbsp;</td>
				<td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
				<td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader" align="center">&nbsp;&nbsp; 
                    <%
If inStr(EntityModificationAccessString,"(8)" ) <> 0 Then
%>
                    <b>Drop</b> 
                    <%
End If
%>
                    &nbsp;&nbsp;</td>
				<td bgcolor="<%=TitleBar%>" width="1%">&nbsp;</td>
				<td bgcolor="<%=TitleBar%>" width="10%" class="CastlesTextHeader" align="center">&nbsp;&nbsp; 
                    <%
If inStr(EntityModificationAccessString,"(8)" ) <> 0 Then
%>
                    <b>Completed</b> 
                    <%
End If
%>
                    &nbsp;&nbsp;</td>
                </tr>
                <%
Response.Flush
if ListingsHasRecords then
'if Len(SQLStmt) <> 0 then
	if Not Listings.EOF then
		SearchResultsReturned = Listings.getrows
		Count = 1
		DeleteCount = 1
		SearchResultsArrayNumRows = uBound(SearchResultsReturned,2)
	
		Field_ID = 0
		Field_DesignerPublishStatusID = 1
		Field_FeaturedProperty = 2
		Field_Address = 3
		Field_Unit = 4
		Field_City = 5
		Field_State = 6
		Field_ZipCode = 7
		Field_ListPrice = 8
		Field_CompanyName = 9
		Field_FirstName = 10
		Field_LastName = 11
		Field_TelNumber = 12
		Field_EmailAddress = 13
		Field_FaxNumber = 14
		Field_PropertyTypeName = 15
		Field_ListingID = 16
		Field_Bedrooms = 17
		Field_FullBaths = 18
		Field_HalfBaths = 19
		Field_LivingArea = 20
		Field_DesignerFirstName = 21
		Field_DesignerLastName = 22
		Field_DesignerCompanyName = 23
		Field_DesignerID = 24
		Field_MoreRecords = 25

		For SearchResultsArrayRowCounter = 0 to SearchResultsArrayNumRows
			DesignerPublishStatusID = SearchResultsReturned(Field_DesignerPublishStatusID,SearchResultsArrayRowCounter)
			FeaturedProperty = SearchResultsReturned(Field_FeaturedProperty,SearchResultsArrayRowCounter)
			Address = SearchResultsReturned(Field_Address,SearchResultsArrayRowCounter)
			Unit = SearchResultsReturned(Field_Unit,SearchResultsArrayRowCounter)
			City = SearchResultsReturned(Field_City,SearchResultsArrayRowCounter)
			State = SearchResultsReturned(Field_State,SearchResultsArrayRowCounter)
			ZipCode = SearchResultsReturned(Field_ZipCode,SearchResultsArrayRowCounter)
			ListPrice = SearchResultsReturned(Field_ListPrice,SearchResultsArrayRowCounter)
			CompanyName = SearchResultsReturned(Field_CompanyName,SearchResultsArrayRowCounter)
			FirstName = SearchResultsReturned(Field_FirstName,SearchResultsArrayRowCounter)
			LastName = SearchResultsReturned(Field_LastName,SearchResultsArrayRowCounter)
			TelNumber = SearchResultsReturned(Field_TelNumber,SearchResultsArrayRowCounter)
			EmailAddress = SearchResultsReturned(Field_EmailAddress,SearchResultsArrayRowCounter)
			FaxNumber = SearchResultsReturned(Field_FaxNumber,SearchResultsArrayRowCounter)
			PropertyTypeName = SearchResultsReturned(Field_PropertyTypeName,SearchResultsArrayRowCounter)
			ListingID = SearchResultsReturned(Field_ListingID,SearchResultsArrayRowCounter)
			Bedrooms = SearchResultsReturned(Field_Bedrooms,SearchResultsArrayRowCounter)
			FullBaths = SearchResultsReturned(Field_FullBaths,SearchResultsArrayRowCounter)
			HalfBaths = SearchResultsReturned(Field_HalfBaths,SearchResultsArrayRowCounter)
			LivingArea = SearchResultsReturned(Field_LivingArea,SearchResultsArrayRowCounter)
			DesignerFirstName = SearchResultsReturned(Field_DesignerFirstName,SearchResultsArrayRowCounter)
			DesignerLastName = SearchResultsReturned(Field_DesignerLastName,SearchResultsArrayRowCounter)
			DesignerCompanyName = SearchResultsReturned(Field_DesignerCompanyName,SearchResultsArrayRowCounter)
			DesignerID = SearchResultsReturned(Field_DesignerID,SearchResultsArrayRowCounter)
			MoreRecords = SearchResultsReturned(Field_MoreRecords,SearchResultsArrayRowCounter)
	
			If Count mod 2 = 0 then
				BgColor = "#FFFFFF"
			Else
				BgColor = LightField 
			End If
%>
                <tr> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td width="1%">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody"><%=LastName%>,&nbsp;<%=FirstName%>&nbsp;<%=MiddleInitial%><br>
                    <%=CompanyName%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody">
				  	<%=Address%>&nbsp;<%=Unit%>
				  </td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="5%" class="CastlesTextBody"> 
                    <%If (DesignerPublishStatusID) = 1 Then%>
                    <img src="/manage/images/icon_active.gif" width="20" height="20" border="0" alt="Active"> 
                    <%Else%>
                    <img src="/manage/images/icon_nonactive.gif" width="20" height="20" border="0" alt="In-Active"> 
                    <%End If%>
                  </td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody"><%=TelNumber%></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="15%" class="CastlesTextBody"><a href="mailto:<%=EmailAddress%>" class="normal"><%=EmailAddress%></a></td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody"> 
                    <%
If inStr(EntityModificationAccessString,"(4)") <> 0 Then
%>
                    <a href="viewadvertisements.asp?ListingID=<%=ListingID%>&SearchListingID=<%=SearchListingID%>&SearchedAddress=<%=Server.URLEncode(SearchedAddress)%>&SearchedSizes=<%=Server.URLEncode(SearchedSizes)%>&SearchedStates=<%=Server.URLEncode(SearchedStates)%>&SearchedStatus=<%=Server.URLEncode(SearchedApartmentStatus)%>&PriceFrom=<%=Server.URLEncode(PriceFrom)%>&PriceTo=<%=Server.URLEncode(PriceTo)%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>" class="normal"><b><%=WordPhrase_View%></b></a>&nbsp;&nbsp;&nbsp;&nbsp; 
                    <%
End If
If inStr(EntityModificationAccessString,"(2)" ) <> 0 Then
%>
                    <a href="editlisting.asp?ListingID=<%=ListingID%>&SearchListingID=<%=SearchListingID%>&SearchedAddress=<%=Server.URLEncode(SearchedAddress)%>&SearchedSizes=<%=Server.URLEncode(SearchedSizes)%>&SearchedStates=<%=Server.URLEncode(SearchedStates)%>&SearchedStatus=<%=Server.URLEncode(SearchedApartmentStatus)%>&PriceFrom=<%=Server.URLEncode(PriceFrom)%>&PriceTo=<%=Server.URLEncode(PriceTo)%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>" class="normal"><b><%=WordPhrase_Edit%></b></a>&nbsp;&nbsp;&nbsp;&nbsp; 
                    <%
End If
If inStr(EntityModificationAccessString,"(6)" ) <> 0 Then
FullName = FirstName & " " & LastName
%>
                    <a href="javascript:onClick=PopNotes('1','<%=BrokerID%>','<%=Server.URLEncode(FullName)%>');" class="normal"><b><%=WordPhrase_Notes%></b></a> 
                    <%
End If
%>
                  </td>
                  <td width="1%" align="center">&nbsp;</td>
                  <td width="11%" class="CastlesTextBody" align="center">&nbsp;&nbsp; 
                    <%
If inStr(EntityModificationAccessString,"(8)" ) <> 0 AND DesignerPublishStatusID <> 2 Then
%>
                    <input type="checkbox" name="PickUpRecordID<%=DeleteCount%>" value="<%=ListingID%>">
<%
else
%>
					Picked up by: <%=DesignerFirstName%>&nbsp;<%=DesignerLastName%>&nbsp;
					<%
					if (DesignerCompanyName <> "") or (not isNull(DesignerCompanyName)) then
					%>
						(<%=DesignerCompanyName%>)
					<%
					end if
					%>
<%
End If
%>
                    &nbsp;&nbsp;</td>
					<td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody" align="center">&nbsp;&nbsp; 
<%					
If (inStr(EntityModificationAccessString,"(8)") <> 0) AND (DesignerPublishStatusID = 2) AND (cStr(DesignerID) = cStr(EntityPrimaryKeyValue)) Then
%>
                    <input type="checkbox" name="DropRecordID<%=DeleteCount%>" value="<%=ListingID%>">
<%
End If
%>
                    &nbsp;&nbsp;</td>
					<td width="1%" align="center">&nbsp;</td>
                  <td width="10%" class="CastlesTextBody" align="center">&nbsp;&nbsp; 
                    <%
If (inStr(EntityModificationAccessString,"(8)") <> 0) AND (DesignerPublishStatusID = 2) AND (cStr(DesignerID) = cStr(EntityPrimaryKeyValue)) Then
%>
                    <input type="checkbox" name="CompletedRecordID<%=DeleteCount%>" value="<%=ListingID%>">
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
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td colspan="18">&nbsp;</td>
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
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td width="1%">&nbsp;</td>
                  <td colspan="18" class="CastlesTextBody"><%=WordPhrase_NoResultsFound%>...</td>
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
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=BgColor%>"> 
                  <td colspan="18">&nbsp;</td>
                </tr>
                <%
		Count = Count+1
	Next
End If
%>
                <tr> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr bgcolor="<%=ThinLine%>"> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="18" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <input type="hidden" name="NumberOfRecordsToPickUp" value="<%=DeleteCount%>">
				<input type="hidden" name="DesignerID" value="<%=EntityPrimaryKeyValue%>">
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
'Generates Appropriate Pagination for Simple System Search Footer
Function DCSystemSimpleSearchPagination(Page,WordPhrase_JumpToPage,PageName,MoreRecords,TotalRecords,RecsPerPage,SearchListingID,SearchedAddress,SearchedSizes,SearchedStates,SearchedApartmentStatus,PriceFrom,PriceTo)
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
			DCSystemSimpleSearchPagination = DCSystemSimpleSearchPagination &  "<a href=" & DQ & PageName & "?ListingID=" & Server.URLEncode(SearchListingID) & "&Address=" & Server.URLEncode(SearchedAddress) & "&Sizes=" & Server.URLEncode(SearchedSizes) & "&States=" & Server.URLEncode(SearchedStates) &  "&Status=" & Server.URLEncode(SearchedApartmentStatus) & "&PriceFrom=" & Server.URLEncode(PriceFrom) & "&PriceTo=" & Server.URLEncode(PriceTo) & "&Page=" & PageCount & "&TotalRecords=" & TotalRecords & DQ &" class=""orange""><b>" & PageCount & "</b></a>|"
		Else
			DCSystemSimpleSearchPagination = DCSystemSimpleSearchPagination & "<b>" & PageCount & "</b>|" 
		End If
		PageCount = PageCount + 1
	Next
End Function					
'Generates Appropriate Pagination Code
if ListingsHasRecords then
'If Len(SQLStmt) <> 0 Then
	PageName = "searchadvertisements.asp"
	TotalRecords = (MoreRecords)+(RecsPerPage)
	PaginationCode = DCSystemSimpleSearchPagination(Page,WordPhrase_JumpToPage,PageName,MoreRecords,TotalRecords,RecsPerPage,SearchListingID,SearchedAddress,SearchedSizes,SearchedStates,SearchedApartmentStatus,PriceFrom,PriceTo)
	Response.Write PaginationCode
End If
%>
                  </td>
                  <td width="50%" align="right" class="CastlesTextBody"> 
                    <%
If inStr(EntityModificationAccessString,"(8)" ) <> 0 Then
%>
                    <input type="submit" name="Submit" value="Change Status" class="CastlesTextBlack">
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
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
    <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
  </tr>
  <tr width="100%"> 
    <td width="100%" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td > 
            <table width="481" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="25">&nbsp;</td>
                <td width="456" class="CastlesTextBody">&copy;&nbsp;<%=DatePart("yyyy",Date)%>&nbsp;Castles Magazine.&nbsp;&nbsp;All rights reserved.&nbsp;&nbsp;.</td>
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
