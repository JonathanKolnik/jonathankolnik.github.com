<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 7
PageTopNavigationSubHeaderID = 33
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

'functions for the dynamic display of rows
Function DisplayFeatures(IDfield, ColumnName, TableName)
	FeatureValue = ListingProfile.Fields.Item(IDfield).Value
	FeatureValue = Replace(FeatureValue,"(","")
	FeatureValue = Replace(FeatureValue,")","")
	FeatureValueArray = Split(FeatureValue,",")
	FeatureValueArrayLength = Ubound(FeatureValueArray)
	Whereclause=""
	'response.write "FeatureValueArrayLength is: " & FeatureValueArrayLength & "<BR>"

	For each x In FeatureValueArray
		if FeatureValueArrayLength = 0 then
			if x <> 0 then
				Whereclause = Whereclause & " " & IDfield & " = " & trim(x)
			end if
		end if
		if FeatureValueArrayLength > 0 then
			if x <> 0 then
				Whereclause = Whereclause & " " & IDfield & " = " & trim(x) & " OR"
			end if
			FeatureValueArrayLength = FeatureValueArrayLength - 1
		end if
	Next
	'response.write "Whereclause is: " & Whereclause & "<BR>"
	set Command1 = Server.CreateObject("ADODB.Command")
	with command1
		.ActiveConnection = Connect
		.CommandText = "Castles_System_DisplayFeatureNames"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.Parameters.Append .CreateParameter("@ColumnName", 200, 1,100,ColumnName)
		.Parameters.Append .CreateParameter("@Tablename", 200, 1,100,Tablename)
		.Parameters.Append .CreateParameter("@Whereclause", 200, 1,1000,Whereclause)
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		set FeatureValueNames = .Execute()
	end with
	set Command1 = nothing
	
	FeatureValueArray = FeatureValueNames.getrows
	FeatureValueNames.close
	set FeatureValueNames = nothing
	FeatureValueArrayNumRows=uBound(FeatureValueArray,2)
	Field_FeatureValueName = 0
	FeatureValue = ""
	For FeatureValueArrayRowCounter = 0 to FeatureValueArrayNumRows
		FeatureValueName = FeatureValueArray(Field_FeatureValueName,FeatureValueArrayRowCounter)
		if FeatureValueArrayNumRows = 0 then
			FeatureValue = FeatureValue & FeatureValueName
		end if
		if FeatureValueArrayNumRows > 0 then
			FeatureValue = FeatureValue & FeatureValueName & ", "
			FeatureValueArrayNumRows = FeatureValueArrayNumRows - 1
		end if
	Next
	Response.write FeatureValue
End Function

Function Exists(fieldname,dbfield,datadesc,datatype)
	fieldvalue = ListingProfile.Fields.Item(dbfield).Value
	dq=""""
	if((fieldvalue<>"")and(fieldvalue<>"N/A")and(fieldvalue<>"0")and(not isNull(fieldvalue))) then
		select case datatype
			case "currency"
				fieldvalue = FormatCurrency(fieldvalue,0)
			case "numbercomma"
				fieldvalue = FormatNumber(fieldvalue,0)
			case "decimal2"
				fieldvalue = Round(fieldvalue,2)
			case else
				fieldvalue = fieldvalue
		end select

		response.write "<tr>" & vbcrlf 
		response.write "    <td class=" &dq& "CastlesTextBody" &dq& "align=" &dq& "left" &dq& "width=" &dq& "100" &dq& ">" & fieldname & ": </td>"&vbcrlf
        response.write "    <td class=" &dq& "CastlesTextBlack" &dq& "width=" &dq& "183" &dq& ">" & fieldvalue & " " & datadesc & "</td>"&vbcrlf
        response.write "</tr>"		
	end if	
End Function

'Listing Profile Info
ListingID = Request.QueryString("ListingID")
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_Listing_Details_View"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@ListingID", 200, 1,200,ListingID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set ListingProfile = .Execute()
End With
Set Command1 = Nothing
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
              <form name="EditEntity" method="post" enctype="multipart/form-data" onSubmit="return Validate();">
                <tr> 
                  <td width="100%" height="1"><img src="/manage/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="100%" bgcolor="<%=LightField%>">
                    <table width="595" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td valign="top" width="300"> 
                          <%
if ListingProfile.Fields.Item("PicturePath1").Value = "" or isNull(ListingProfile.Fields.Item("PicturePath1").Value) then
	PicturePath = "http://castlesmag.com/manage/images/noPic.gif"
else
	PicturePath = ListingProfile.Fields.Item("PicturePath1").Value
end if
if ListingProfile.Fields.Item("PicturePath2").Value = "" or isNull(ListingProfile.Fields.Item("PicturePath2").Value) then
	morePics = false
else
	morePics = true
end if
%>
                          <table width="300" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="300" colspan="2" bgcolor="#EDEBDB"><img src="<%=PicturePath%>" name="ListingPic" width="300" height="300"></td>
                            </tr>
                            <tr> 
                              <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                              <td width="290" bgcolor="#EDEBDB" class="CastlesTextBody"> 
                                <%
if morePics then
%>
                                More images: <a href="javascript:switchImage(1)" class="normal">1</a> 
                                | <a href="javascript:switchImage(2)" class="normal">2</a> 
                                <%	
else
	response.write "&nbsp;"
end if
%>
                              </td>
                            </tr>
                            <tr> 
                              <td colspan="2" height="1"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
                            </tr>
                            <tr> 
                              <td colspan="2" height="1" bgcolor="#CCCCCC"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
                            </tr>
                            <tr> 
                              <td colspan="2" height="10"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                            </tr>
							<tr> 
                              <td colspan="2" align="left" class="CastlesTextBody" bgcolor="#EDEBDB">This 
                                Property Listing is provided by:</td>
                            </tr>
							<tr>
								<td colspan="2" align="left" class="CastlesTextBody">&nbsp;</td>
							</tr>
                            <%
CompanyName = ListingProfile.Fields.Item("CompanyName").Value
LastName = ListingProfile.Fields.Item("LastName").Value
FirstName = ListingProfile.Fields.Item("FirstName").Value
MiddleInitial = ListingProfile.Fields.Item("MiddleInitial").Value
BrokerAddressLine1 = ListingProfile.Fields.Item("BrokerAddressLine1").Value
BrokerAddressLine2 = ListingProfile.Fields.Item("BrokerAddressLine2").Value
BrokerCity = ListingProfile.Fields.Item("BrokerCity").Value
TelNumber = ListingProfile.Fields.Item("TelNumber").Value
FaxNumber = ListingProfile.Fields.Item("FaxNumber").Value
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody" valign="top">Company: 
                              </td>
                              <td width="183" class="CastlesTextBlack"><%=CompanyName%></td>
                            </tr>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody" valign="top">Agent: 
                              </td>
                              <td width="183" class="CastlesTextBlack"><%=LastName%>, <%=FirstName%> <%=MiddleInitial%></td>
                            </tr>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody" valign="top">Address:</td>
                              <td width="183" class="CastlesTextBlack"><%=BrokerAddressLine1%><br>
                                <% if (Len(BrokerAddressLine2)<>0) then
					  		response.write BrokerAddressLine2 & "<BR>"
						end if
					  %>
                                <%=BrokerCity%><br>
                                Tel: <%=TelNumber%><br>
                                <% if (Len(FaxNumber)<>0) then
					  		response.write "Fax:  " & FaxNumber & "<BR>"
						end if
					  %>
                              </td>
                            </tr>
                          </table>
                        </td>
                        <td valign="top" width="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                        <td valign="top" width="1" bgcolor="#CCCCCC"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                        <td valign="top" width="10">&nbsp;</td>
                        <td valign="top" width="283"> 
                          <table width="283" border="0" cellspacing="0" cellpadding="2">
                            <tr> 
                              <td class="CastlesTextBody" align="left" width="150"> 

                                <b><%=ListingProfile.Fields.Item("Address").Value%>&nbsp;<%=ListingProfile.Fields.Item("Unit").Value%> </b> 

                              </td>
                              <td class="CastlesTextBody" align="right"><b><%=FormatCurrency(ListingProfile.Fields.Item("ListPrice").Value,0)%></b></td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody" align="left" width="150">&nbsp;</td>
                              <td class="CastlesTextBody" align="right">&nbsp;</td>
                            </tr>
                          </table>
                          <table width="283" border="0" cellspacing="0" cellpadding="2">
			    <tr> 
                              <td width="100" align="left" class="CastlesTextBody">ListingID:</td>
                              <td width="183" class="CastlesTextBlack"><%=ListingID%> </td>
                            </tr>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Address: 
                              </td>
                              <td width="183" class="CastlesTextBlack"><%=ListingProfile.Fields.Item("City").Value%>,&nbsp;<%=ListingProfile.Fields.Item("StateProvinceName").Value%>&nbsp;<%=ListingProfile.Fields.Item("Zipcode").Value%><br>
                                <%=ListingProfile.Fields.Item("CountryName").Value%> </td>
                            </tr>
                            <%
Exists "Bedrooms","Bedrooms","",""
Exists "Full Baths","FullBaths","",""
Exists "Half Baths","HalfBaths","",""
Exists "Interior Description","InteriorDescription","",""
Exists "Exterior Description","ExteriorDescription","",""
Exists "Property Type","PropertyTypeName","",""
Exists "Property Style","PropertyStyleName","",""
Exists "Living Area","LivingArea","",""
Exists "Exterior Type","ExteriorTypeName","",""
Exists "Waterfront","Waterfront","",""
Exists "Ski","Ski","",""
Exists "Condo","Condo","",""
Exists "Resort","Resort","",""
Exists "Country Club","CountryClub","",""
Exists "Farm/Ranch","FarmOrRanch","",""
Exists "Castle","Castle","",""
Exists "Assessment","Assessment","","currency"
Exists "Taxes","Taxes","","currency"
Exists "Tax Year","TaxYear","",""

if ((ListingProfile.Fields.Item("InteriorFeatureID").Value <> "")and(ListingProfile.Fields.Item("InteriorFeatureID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Interior 
                                Features: </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "InteriorFeatureID","InteriorFeatureName","InteriorFeatures"
									%>
                              </td>
                            </tr>
                            <%
end if
if ((ListingProfile.Fields.Item("ExteriorFeatureID").Value <> "")and(ListingProfile.Fields.Item("ExteriorFeatureID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Exterior 
                                Features: </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "ExteriorFeatureID","ExteriorFeatureName","ExteriorFeatures"
									%>
                              </td>
                            </tr>
                            <%
end if
if ((ListingProfile.Fields.Item("BasementFeatureID").Value <> "")and(ListingProfile.Fields.Item("BasementFeatureID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Basement 
                                Features: </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "BasementFeatureID","BasementFeatureName","BasementFeatures"
									%>
                              </td>
                            </tr>
                            <%
end if
if ((ListingProfile.Fields.Item("FlooringID").Value <> "")and(ListingProfile.Fields.Item("FlooringID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Flooring: 
                              </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "FlooringID","FlooringName","Flooring"
									%>
                              </td>
                            </tr>
                            <%
end if
if ((ListingProfile.Fields.Item("HeatingID").Value <> "")and(ListingProfile.Fields.Item("HeatingID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Basement 
                                Features: </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "HeatingID","HeatingName","Heating"
									%>
                              </td>
                            </tr>
                            <%
end if
if ((ListingProfile.Fields.Item("CoolingID").Value <> "")and(ListingProfile.Fields.Item("CoolingID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Cooling: 
                              </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "CoolingID","CoolingName","Cooling"
									%>
                              </td>
                            </tr>
                            <%
end if
if ((ListingProfile.Fields.Item("HotWaterID").Value <> "")and(ListingProfile.Fields.Item("HotWaterID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Hot 
                                Water: </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "HotWaterID","HotWaterName","HotWater"
									%>
                              </td>
                            </tr>
                            <%
end if
if ((ListingProfile.Fields.Item("ApplianceID").Value <> "")and(ListingProfile.Fields.Item("ApplianceID").Value <> "(0)")) then
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">Appliances: 
                              </td>
                              <td width="183" class="CastlesTextBlack"> 
                                <%
									DisplayFeatures "ApplianceID","ApplianceName","Appliances"
									%>
                              </td>
                            </tr>
                            <%
end if
%>
                            <tr> 
                              <td width="100" align="left" class="CastlesTextBody">&nbsp;</td>
                              <td width="183" class="CastlesTextBlack">&nbsp;</td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                    <!--here-->
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
                <td width="275" align="right"></td>
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
