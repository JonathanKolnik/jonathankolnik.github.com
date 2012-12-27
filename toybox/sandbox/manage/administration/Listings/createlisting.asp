<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 7
PageTopNavigationSubHeaderID = 31
FromPageTopNavigationSubHeaderID = 29
EntityID = 10
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
WordPhrase_BrokerInfo = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_DirectLine = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_EditBroker = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
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
SearchListingID = Request.QueryString("SearchListingID")
SearchedAddress = Request.QueryString("SearchedAddress")
SearchedSizes = Request.QueryString("SearchedSizes")
SearchedAreas = Request.QueryString("SearchedAreas")
SearchedStatus = Request.QueryString("SearchedStatus")
PriceFrom = Request.QueryString("PriceFrom")
PriceTo = Request.QueryString("PriceTo")

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
function IsNumeric(field,charset){
	//  check for valid numeric strings
	// charset 1 for money values
	if(charset == 1){	
		var ValidChars = "0123456789.";
	}
	
	else{
		var ValidChars = "0123456789";
	}
	var Char;
	var Result = true;
	if (field.length == 0){
		return false;
	}
	//  test field consists of valid characters listed above
	for (i = 0; i < field.length && Result == true; i++){
		Char = field.charAt(i);
		if (ValidChars.indexOf(Char) == -1){
			Result = false;
		}
	}
	return Result;
}
function Validate(){
	var errorString = "";
	var errorTrue = "";

	if (document.EditEntity.Address.value == "") {
		errorString=errorString + " - Please enter the Address. \r";
		errorTrue="y";
	}
	if (document.EditEntity.City.value == "") {
		errorString=errorString + " - Please enter the City. \r";
		errorTrue="y";
	}
	if (document.EditEntity.StateProvinceID.value == "") {
		errorString=errorString + " - Please enter the State. \r";
		errorTrue="y";
	}
	if (document.EditEntity.ZipCode.value == "") {
		errorString=errorString + " - Please enter the Zipcode. \r";
		errorTrue="y";
	}
	if (document.EditEntity.ZipCode.value != "") {
		if (IsNumeric(document.EditEntity.ZipCode.value,2) == false) {
			errorString=errorString + " - Zipcode may only contain numbers. \r";
			errorTrue="y";
		}
	}	
	/*if (document.EditEntity.TotalRooms.value == "") {
		errorString=errorString + " - Please enter the number of Total Rooms. \r";
		errorTrue="y";
	}*/
	if (document.EditEntity.TotalRooms.value != "") {
		if (IsNumeric(document.EditEntity.TotalRooms.value,2) == false) {
			errorString=errorString + " - Total Rooms may only contain numbers. \r";
			errorTrue="y";
		}
	}	
	/*if (document.EditEntity.Bedrooms.value == "") {
		errorString=errorString + " - Please enter the number of Bedrooms. \r";
		errorTrue="y";
	}*/
	if (document.EditEntity.Bedrooms.value != "") {
		if (IsNumeric(document.EditEntity.Bedrooms.value,2) == false) {
			errorString=errorString + " - Bedrooms may only contain numbers. \r";
			errorTrue="y";
		}
	}	
	/*if (document.EditEntity.FullBaths.value == "") {
		errorString=errorString + " - Please enter the number of Full Baths. \r";
		errorTrue="y";
	}*/
	if (document.EditEntity.FullBaths.value != "") {
		if (IsNumeric(document.EditEntity.FullBaths.value,2) == false) {
			errorString=errorString + " - Full Baths may only contain numbers. \r";
			errorTrue="y";
		}
	}
	
	/* 	New checks added by Piyush Kacha on 15 Mar 2005 */
	/***************/
	if (document.EditEntity.HalfBaths.value != "") {
		if (IsNumeric(document.EditEntity.HalfBaths.value,2) == false) {
			errorString=errorString + " - Half Baths may only contain numbers. \r";
			errorTrue="y";
		}
	}

	if (document.EditEntity.LotSize.value != "") {
			if (IsNumeric(document.EditEntity.LotSize.value,2) == false) {
				errorString=errorString + " - Lot Size may only contain numbers. \r";
				errorTrue="y";
			}
		}

	if (document.EditEntity.YearBuilt.value != "") {
			if (IsNumeric(document.EditEntity.YearBuilt.value,2) == false) {
				errorString=errorString + " - Year Built may only contain numbers. \r";
				errorTrue="y";
			}
		}
	/*******************/
	/*if (document.EditEntity.LivingArea.value == "") {
		errorString=errorString + " - Please enter the Living Area. \r";
		errorTrue="y";
	}*/
	if (document.EditEntity.LivingArea.value != "") {
		if (IsNumeric(document.EditEntity.LivingArea.value,2) == false) {
			errorString=errorString + " - Living Area may only contain numbers. \r";
			errorTrue="y";
		}
	}
	
	if (document.EditEntity.ListPrice.value == "") {
		errorString=errorString + " - Please enter the List Price. \r";
		errorTrue="y";
	}

	if (document.EditEntity.ListPrice.value != "") {
		if (IsNumeric(document.EditEntity.ListPrice.value,1) == false) {
			errorString=errorString + " - List Price may only contain numbers. \r";
			errorTrue="y";
		}
	}	
	/*if (document.EditEntity.Assessment.value == "") {
		errorString=errorString + " - Please enter the Assessment. \r";
		errorTrue="y";
	}*/

	if (document.EditEntity.Assessment.value != "") {
		if (IsNumeric(document.EditEntity.Assessment.value,1) == false) {
			errorString=errorString + " - Assessment may only contain numbers. \r";
			errorTrue="y";
		}
	}	
	/*if (document.EditEntity.TaxYear.value == "") {
		errorString=errorString + " - Please enter the Tax Year. \r";
		errorTrue="y";
	}*/
	if (document.EditEntity.TaxYear.value != "") {
		if (IsNumeric(document.EditEntity.TaxYear.value,2) == false) {
			errorString=errorString + " - Tax Year may only contain numbers. \r";
			errorTrue="y";
		}
	}
	/*if (document.EditEntity.Taxes.value == "") {
		errorString=errorString + " - Please enter the Annual Real Estate Taxes. \r";
		errorTrue="y";
	}*/
	if (document.EditEntity.Taxes.value != "") {
		if (IsNumeric(document.EditEntity.Taxes.value,1) == false) {
			errorString=errorString + " - Annual Real Estate Taxes may only contain numbers. \r";
			errorTrue="y";
		}
	}	
	
	if (errorTrue == "y") {
		alert("Missing Required Fields: \r" + errorString);
		return false;
	}else {
		return true;
	}
}
function GetState() {
if (document.EditEntity.CountryID.options(document.EditEntity.CountryID.options.selectedIndex).value != 0 && document.EditEntity.CountryID.options(document.EditEntity.CountryID.options.selectedIndex).value != 1)
{
document.EditEntity.StateProvinceID.options(0).selected = true;
document.EditEntity.StateProvinceID.disabled = true;
}
else
{
document.EditEntity.StateProvinceID.disabled = false;
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
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#000000" width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td bgcolor="<%=TitleBar%>" width="100%" class="CastlesTextHeader" height="20">&nbsp;&nbsp;&nbsp;<b>LISTING INFORMATION</b></td>
              </tr>
              <form name="EditEntity" method="post" enctype="multipart/form-data" onSubmit="return Validate()" action="drivelistings.asp?DCDataDriverType=SQLInsert&SearchListingID=<%=SearchListingID%>&SearchedAddress=<%=SearchedAddress%>&SearchedSizes=<%=SearchedSizes%>&SearchedAreas=<%=SearchedAreas%>&SearchedStatus=<%=SearchedStatus%>&PriceFrom=<%=PriceFrom%>&PriceTo=<%=PriceTo%>&Page=<%=Page%>&TotalRecords=<%=TotalRecords%>">
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
					  			<td class="CastlesTextBody"><font color="#FF0000">* - Do not use commas,$ signs or any special characters</font></td>
					  		</tr>	
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;1&nbsp;] Address:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="Address" class="CastlesTextBlack" size="30" maxlength="100">
                                &nbsp;&nbsp; </td>
                            </tr>							
							<tr> 
                              <td class="CastlesTextBody"><input type="checkbox" name="ShowAddress" value="Y">&nbsp;Show Address on Website?</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;2&nbsp;] Unit:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="Unit" class="CastlesTextBlack" size="10" maxlength="20">
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;3&nbsp;] City:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="City" class="CastlesTextBlack" size="30" maxlength="100">
                              </td>
                            </tr>
                            <tr> 
                              <td width="200" class="CastlesTextBody">[&nbsp;4&nbsp;] 
                                State: </td>
                            </tr>
                            <tr> 
                              <td width="200" class="CastlesTextBlack"> 
                                <select name="StateProvinceID" class="CastlesTextBlack" onchange="GetState();">
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_StateProvince_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set StateProvince = .Execute()
End With
Set Command1 = Nothing

While (NOT StateProvince.EOF)

%>
                                  <option value="<%=(StateProvince.Fields.Item("StateProvinceID").Value)%>"><%=(StateProvince.Fields.Item("StateProvinceName").Value)%></option>
<%
	StateProvince.MoveNext()
Wend
StateProvince.close
Set StateProvince = Nothing
%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td width="200" class="CastlesTextBody">[&nbsp;5&nbsp;] 
                                Zip/Postal Code: </td>
                            </tr>
                            <tr> 
                              <td width="200" class="CastlesTextBlack"> 
                                <input type="text" name="ZipCode" class="CastlesTextBlack" size="20" maxlength="15"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
							<tr> 
                              <td width="200" class="CastlesTextBody">[&nbsp;6&nbsp;] 
                                Country: </td>
                            </tr>
                            <tr> 
                              <td width="200" class="CastlesTextBlack"> 
                                <select name="CountryID" class="CastlesTextBlack" onchange="GetState();">
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_Country_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Country = .Execute()
End With
Set Command1 = Nothing

While (NOT Country.EOF)
%>
                                  <option value="<%=(Country.Fields.Item("CountryID").Value)%>"><%=(Country.Fields.Item("CountryName").Value)%></option>
<%
	Country.MoveNext()
Wend
Country.close
Set Country = Nothing
%>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;7&nbsp;] Area 
                                Type:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="AreaTypeID" class="CastlesTextBlack">
<%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_AreaTypes_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set AreaTypes = .Execute()
End With
Set Command1 = Nothing

While (NOT AreaTypes.EOF)

%>

                           <option value="<%=(AreaTypes.Fields.Item("AreaTypeID").Value)%>"><%=(AreaTypes.Fields.Item("AreaTypeName").Value)%></option>
<%
	AreaTypes.MoveNext()
Wend
AreaTypes.close
Set AreaTypes = Nothing
%>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;8&nbsp;] Property 
                                Style:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="PropertyStyleID" class="CastlesTextBlack">
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_PropertyStyles_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set PropertyStyles = .Execute()
End With
Set Command1 = Nothing

While (NOT PropertyStyles.EOF)
%>

                 <option value="<%=(PropertyStyles.Fields.Item("PropertyStyleID").Value)%>"><%=(PropertyStyles.Fields.Item("PropertyStyleName").Value)%></option>
<%
	PropertyStyles.MoveNext()
Wend
PropertyStyles.close
Set PropertyStyles = Nothing
%>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;9&nbsp;] Property Type:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="PropertyTypeID" class="CastlesTextBlack">
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_PropertyTypes_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set PropertyTypes = .Execute()
End With
Set Command1 = Nothing

While (NOT PropertyTypes.EOF)
	
%>

              <option value="<%=(PropertyTypes.Fields.Item("PropertyTypeID").Value)%>"><%=(PropertyTypes.Fields.Item("PropertyTypeName").Value)%></option>
<%
	PropertyTypes.MoveNext()
Wend
PropertyTypes.close
Set PropertyTypes = Nothing
%>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;10&nbsp;] Total 
                                Rooms:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="TotalRooms" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;11&nbsp;] Bedrooms:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="Bedrooms" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;12&nbsp;] Full Baths:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="FullBaths" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;13&nbsp;] Half 
                                Baths:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="HalfBaths" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;14&nbsp;] Lot 
                                Size:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
								<input type="text" name="LotSize" class="CastlesTextBlack" size="10" maxlength="20">
                                sq. ft. <font color="#FF0000">*</font></td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;15&nbsp;] Year 
                                Built:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
								<input type="text" name="YearBuilt" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span></td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;16&nbsp;] Living 
                                Area:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
								<input type="text" name="LivingArea" class="CastlesTextBlack" size="10" maxlength="20">
                                sq. ft. <font color="#FF0000">*</font></td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;17&nbsp;] List 
                                Price:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="ListPrice" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody"><input type="checkbox" name="ShowListPrice" value="Y">&nbsp;Show List Price on Website?</td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;18&nbsp;] Assessment:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="Assessment" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;19&nbsp;] Tax 
                                Year:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
								<input type="text" name="TaxYear" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;20&nbsp;] Annual 
                                Real Estate Taxes:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
								<input type="text" name="Taxes" class="CastlesTextBlack" size="10" maxlength="20"> <span class="CastlesTextBody"><font color="#FF0000">*</font></span>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;21&nbsp;] Waterfront:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="Waterfront" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;22&nbsp;] Ski:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="Ski" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;23&nbsp;] Condo:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="Condo" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;24&nbsp;] Resort:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="Resort" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;25&nbsp;] Country 
                                Club:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="CountryClub" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;26&nbsp;] Farm/Ranch:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="FarmOrRanch" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;26B&nbsp;] Castle:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="Castle" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N" selected><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;27&nbsp;] <%=WordPhrase_Active%>:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="Active" class="CastlesTextBlack">
                                  <option value="Y" selected><%=WordPhrase_Yes%></option>
                                  <option value="N"><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;28&nbsp;] Featured:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <select name="FeaturedProperty" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N"><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
                            <tr>
                              <td class="CastlesTextBody">[&nbsp;28B&nbsp;] Current Listing:</td>
                            </tr>
                            <tr>
                              <td class="CastlesTextBlack">
                                <select name="CurrentListing" class="CastlesTextBlack">
                                  <option value="Y"><%=WordPhrase_Yes%></option>
                                  <option value="N"><%=WordPhrase_No%></option>
                                </select>
                              </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;29&nbsp;] Featured 
                                Title:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="FeaturedTitle" class="CastlesTextBlack" size="30" maxlength="100">
                                &nbsp;&nbsp; </td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;30&nbsp;] Featured 
                                Description:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <textarea name="FeaturedDescription" cols="30" rows="5" class="CastlesTextBlack" wrap="VIRTUAL"></textarea>
                              </td>
                            </tr>
                          </table>
                        </td>
                        <td width="365" class="CastlesTextBody" valign="top"> 
                          <table cellspacing="0" cellpadding="3" border="0" width="365">
                            <tr> 
                              <td  class="CastlesTextBody">[&nbsp;31&nbsp;] Hot 
                                Water:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="HotWaterID" class="CastlesTextBlack" multiple size=5>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_HotWater_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set HotWater = .Execute()
End With
Set Command1 = Nothing

While (NOT HotWater.EOF)	
%>

                        <option value="(<%=(HotWater.Fields.Item("HotWaterID").Value)%>)"><%=(HotWater.Fields.Item("HotWaterName").Value)%></option>
<%
	HotWater.MoveNext()
Wend
HotWater.close
Set HotWater = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;32&nbsp;] Heating:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="HeatingID" class="CastlesTextBlack" multiple size=5>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_Heating_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Heating = .Execute()
End With
Set Command1 = Nothing

While (NOT Heating.EOF)
	
%>

                    <option value="(<%=(Heating.Fields.Item("HeatingID").Value)%>)"><%=(Heating.Fields.Item("HeatingName").Value)%></option>
<%
	Heating.MoveNext()
Wend
Heating.close
Set Heating = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;33&nbsp;] Cooling:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="CoolingID" class="CastlesTextBlack" multiple size=5>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_Cooling_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Cooling = .Execute()
End With
Set Command1 = Nothing

While (NOT Cooling.EOF)
	
%>

                      <option value="(<%=(Cooling.Fields.Item("CoolingID").Value)%>)"><%=(Cooling.Fields.Item("CoolingName").Value)%></option>
<%
	Cooling.MoveNext()
Wend
Cooling.close
Set Cooling = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;34&nbsp;] Exterior 
                                Type:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="ExteriorTypeID" class="CastlesTextBlack">
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_ExteriorTypes_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set ExteriorTypes = .Execute()
End With
Set Command1 = Nothing

While (NOT ExteriorTypes.EOF)

%>

                      <option value="<%=(ExteriorTypes.Fields.Item("ExteriorTypeID").Value)%>"><%=(ExteriorTypes.Fields.Item("ExteriorTypeName").Value)%></option>
<%
	ExteriorTypes.MoveNext()
Wend
ExteriorTypes.close
Set ExteriorTypes = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;35&nbsp;] Exterior 
                                Features:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="ExteriorFeatureID" class="CastlesTextBlack" multiple size=5>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_ExteriorFeatures_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set ExteriorFeatures = .Execute()
End With
Set Command1 = Nothing

While (NOT ExteriorFeatures.EOF)

%>

                   <option value="(<%=(ExteriorFeatures.Fields.Item("ExteriorFeatureID").Value)%>)"><%=(ExteriorFeatures.Fields.Item("ExteriorFeatureName").Value)%></option>
                                  <%
	ExteriorFeatures.MoveNext()
Wend
ExteriorFeatures.close
Set ExteriorFeatures = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;36&nbsp;] Interior 
                                Features:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="InteriorFeatureID" class="CastlesTextBlack" multiple size=5>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_InteriorFeatures_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set InteriorFeatures = .Execute()
End With
Set Command1 = Nothing

While (NOT InteriorFeatures.EOF)
%>

                          <option value="(<%=(InteriorFeatures.Fields.Item("InteriorFeatureID").Value)%>)"><%=(InteriorFeatures.Fields.Item("InteriorFeatureName").Value)%></option>
<%
	InteriorFeatures.MoveNext()
Wend
InteriorFeatures.close
Set InteriorFeatures = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;37&nbsp;] Appliances:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="ApplianceID" class="CastlesTextBlack" multiple size=5>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_Appliances_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Appliances = .Execute()
End With
Set Command1 = Nothing

While (NOT Appliances.EOF)
%>

                      <option value="(<%=(Appliances.Fields.Item("ApplianceID").Value)%>)"><%=(Appliances.Fields.Item("ApplianceName").Value)%></option>
<%
	Appliances.MoveNext()
Wend
Appliances.close
Set Appliances = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;38&nbsp;] Flooring:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="FlooringID" class="CastlesTextBlack" multiple size=5>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_Flooring_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Flooring = .Execute()
End With
Set Command1 = Nothing

While (NOT Flooring.EOF)
%>

                                  <option value="(<%=(Flooring.Fields.Item("FlooringID").Value)%>)"><%=(Flooring.Fields.Item("FlooringName").Value)%></option>
                                  <%
	Flooring.MoveNext()
Wend
Flooring.close
Set Flooring = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;39&nbsp;] Foundation 
                                Type:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="FoundationTypeID" class="CastlesTextBlack">
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_FoundationTypes_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set FoundationTypes = .Execute()
End With
Set Command1 = Nothing

While (NOT FoundationTypes.EOF)

%>

                                  <option value="<%=(FoundationTypes.Fields.Item("FoundationTypeID").Value)%>"><%=(FoundationTypes.Fields.Item("FoundationTypeName").Value)%></option>
                                  <%
	FoundationTypes.MoveNext()
Wend
FoundationTypes.close
Set FoundationTypes = Nothing
%>
                                </select>
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;40&nbsp;] Interior 
                                Description:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <textarea name="InteriorDescription" cols="30" rows="5" class="CastlesTextBlack" wrap="VIRTUAL"></textarea>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;41&nbsp;] Exterior 
                                Description:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <textarea name="ExteriorDescription" cols="30" rows="5" class="CastlesTextBlack" wrap="VIRTUAL"></textarea>
                              </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;42&nbsp;] Broker:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody"> 
                                <select name="BrokerID" class="CastlesTextBlack">
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_Brokers_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Brokers = .Execute()
End With
Set Command1 = Nothing

While (NOT Brokers.EOF)

%>

                                  <option value="<%=(Brokers.Fields.Item("BrokerID").Value)%>"><%=(Brokers.Fields.Item("CompanyName").Value)%> 
                                  - <%=(Brokers.Fields.Item("LastName").Value)%>, 
                                  <%=(Brokers.Fields.Item("FirstName").Value)%></option>
<%
	Brokers.MoveNext()
Wend
Brokers.close
Set Brokers = Nothing
%>
                                </select>
                              </td>
                            </tr>
<%
fieldNumber = 43
For i = 1 to 8	
%>							
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;<%=fieldNumber%>&nbsp;] Picture 
                                <%=i%> Preview:</td>
                            </tr>
							<tr>
								<td class="CastlesTextBody"><font color="#FF0000">
									<b>Note:</b> Your images need to be 300 pixels x 300 pixels.</font>
								</td>
							</tr>
                            <tr valign="top">
<%
fieldNumber = fieldNumber + 1
Pic = "http://castlesmag.com/manage/images/noPic.gif"
%>
                              <td width="65%" class="CastlesTextBody"><img src="<%=Pic%>" width="300" height="300"></td>
                            </tr>
							<tr> 
                              <td class="CastlesTextBody">[&nbsp;<%=fieldNumber%>&nbsp;] Picture 
                                <%=i%> Dimensions (width x height):</td>
                            </tr>
                            <tr> 
                              <td width="65%" class="CastlesTextBlack"> 
                                <select name="PictureWidth<%=i%>" class="CastlesTextBlack">
                                  <option value="300" selected>300px</option>
                                </select>
                                x 
                                <select name="PictureHeight<%=i%>" class="CastlesTextBlack">
                                  <option value="300" selected>300px</option>
                                </select>
                              </td>
                            </tr>
<%
	fieldNumber = fieldNumber + 1
%>							
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;<%=fieldNumber%>&nbsp;] Picture 
                                <%=i%> Path:</td>
                            </tr>
                            <tr> 
                              <td width="65%" class="DC_RealEstateNormalText"> 
                                <input type="FILE" name="PicturePath<%=i%>" size="40" class="CastlesTextBlack">
                              </td>
                            </tr>
<%
	fieldNumber = fieldNumber + 1
Next
%>		
							<tr> 
                              <td class="CastlesTextBody">[ Publish ]:</td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">
							  <input type="checkbox" name="ListingPublishStatusID" value="2" class="CastlesTextBlack">
                                Publish to this month's magazine?</td>
                            </tr>					
                          </table>
                        </td>
                      </tr>
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td rowspan="13" valign="top"> <br>
                          <table width="300" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                
                                <input type="submit" name="Submit" value="Create Listing" class="CastlesTextBlack">
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
