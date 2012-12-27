<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 2
PageTopNavigationSubHeaderID = 23
EntityID = 6
%>
<!--#include virtual="/templates/Castlessystemcnekt.asp" -->
<%
'---------------Begin Page Level Multilingual Translation-----------------------
WordPhrasesOnPage = "Edit WebSite Content(|)Languages(|)WebSite Content(|)WebSite Content Body(|)WebSite Content Caption(|)WebSite Content Caption Header"
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

WordPhrase_EditWebSiteContent = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_Languages = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_WebSiteContent = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_WebSiteContentBody = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_WebSiteContentCaption = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
TranslateCount = TranslateCount + 1
WordPhrase_WebSiteContentCaptionHeader = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)

'---------------End Multilingual Translation-----------------------

'Persist Search Values
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_WebSiteContactInfo"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set WebSiteContactInfo = .Execute()
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
function Validate(){
	var errorString = ""
	var errorTrue = ""

	//if (document.EditEntity.WebSiteContentCaption1.value.length > 2000) {
	//	errorString=errorString + "You are only allowed 2000 characters for this field. \r"
	//	errorTrue="y"
	//}
	if (errorTrue == "y") {
		alert(errorString) 
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
          <td bgcolor="<%=ThinLine%>"><img src="../../../../Castles/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#000000" width="100%" height="1"><img src="../../../../Castles/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td bgcolor="<%=TitleBar%>" width="100%" class="CastlesTextHeader" height="20">&nbsp;&nbsp;&nbsp;<b>WebSite 
                  Contact Info</b></td>
              </tr>
              <form name="EditEntity" method="post" onSubmit="return Validate();" action="drivewebsitecontactinfo.asp?DCDataDriverType=SQLUpdate">
                <tr> 
                  <td width="100%" height="1"><img src="../../../../Castles/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="100%" bgcolor="<%=LightField%>"> 
                    <table width="750" border="0" cellspacing="0" cellpadding="2">
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td rowspan="14" valign="top"> 
                          <table width="750" border="0" cellspacing="0" cellpadding="2">
                            <tr> 
                              <td width="5">&nbsp;</td>
                              <td rowspan="14" valign="top"> 
                                <table width="350" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td class="CastlesTextBody">[&nbsp;1&nbsp;] 
                                      Company Name:</td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBlack"> 
                                      <input type="text" name="CompanyName" class="CastlesTextBlack" size="25" maxlength="50" value="<%=WebSiteContactInfo.Fields.Item("CompanyName").Value%>">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBody">[&nbsp;2&nbsp;] 
                                      Address:</td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBlack"> 
                                      <table width="350" border="0" cellspacing="0" cellpadding="2">
                                        <tr> 
                                          <td width="100" align="right" class="CastlesTextBody">Line 
                                            1:</td>
                                          <td width="250" class="CastlesTextBlack"> 
                                            <input type="text" name="AddressLine1" class="CastlesTextBlack" size="25" maxlength="100" value="<%=WebSiteContactInfo.Fields.Item("AddressLine1").Value%>">
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="100" align="right" class="CastlesTextBody">Line 
                                            2:</td>
                                          <td width="250" class="CastlesTextBlack"> 
                                            <input type="text" name="AddressLine2" class="CastlesTextBlack" size="15" maxlength="50" value="<%=WebSiteContactInfo.Fields.Item("AddressLine2").Value%>">
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="100" align="right" class="CastlesTextBody">City:</td>
                                          <td width="250" class="CastlesTextBlack"> 
                                            <input type="text" name="City" class="CastlesTextBlack" size="25" maxlength="100" value="<%=WebSiteContactInfo.Fields.Item("City").Value%>">
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="100" align="right" class="CastlesTextBody">State:</td>
                                          <td width="250" class="CastlesTextBlack"> 
                                            <select name="StateProvinceID" class="CastlesTextBlack">
                                              <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_StateProvince_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set StateProvinces = .Execute()
End With
Set Command1 = Nothing

While (NOT StateProvinces.EOF)
	if WebSiteContactInfo.Fields.Item("StateProvinceID").Value = StateProvinces.Fields.Item("StateProvinceID").Value then
%>
                                              <option value="<%=(StateProvinces.Fields.Item("StateProvinceID").Value)%>" selected><%=(StateProvinces.Fields.Item("StateProvinceName").Value)%></option>
                                              <%
	Else
%>
                                              <option value="<%=(StateProvinces.Fields.Item("StateProvinceID").Value)%>"><%=(StateProvinces.Fields.Item("StateProvinceName").Value)%></option>
                                              <%
	End If
	StateProvinces.MoveNext()
Wend
StateProvinces.close
Set StateProvinces = Nothing
%>
                                            </select>
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="100" align="right" class="CastlesTextBody">Zip 
                                            Code:</td>
                                          <td width="250" class="CastlesTextBlack"> 
                                            <input type="text" name="ZipPostalCode" class="CastlesTextBlack" size="15" maxlength="50" value="<%=WebSiteContactInfo.Fields.Item("ZipPostalCode").Value%>">
                                          </td>
                                        </tr>
                                        <tr> 
                                          <td width="100" align="right" class="CastlesTextBody">Country:</td>
                                          <td width="250" class="CastlesTextBlack"> 
                                            <select name="CountryID" class="CastlesTextBlack">
                                              <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_System_Country_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Countries = .Execute()
End With
Set Command1 = Nothing

While (NOT Countries.EOF)
	if WebSiteContactInfo.Fields.Item("CountryID").Value = Countries.Fields.Item("CountryID").Value then
%>
                                              <option value="<%=(Countries.Fields.Item("CountryID").Value)%>" selected><%=(Countries.Fields.Item("CountryName").Value)%></option>
                                              <%
	Else
%>
                                              <option value="<%=(Countries.Fields.Item("CountryID").Value)%>"><%=(Countries.Fields.Item("CountryName").Value)%></option>
                                              <%
	End If
	Countries.MoveNext()
Wend
Countries.close
Set Countries = Nothing
%>
                                            </select>
                                          </td>
                                        </tr>
                                      </table>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBody">[&nbsp;3&nbsp;] 
                                      Main Phone Number:</td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBlack"> 
                                      <input type="text" name="MainTelNumber" class="CastlesTextBlack" size="15" maxlength="50" value="<%=WebSiteContactInfo.Fields.Item("MainTelNumber").Value%>">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBody">[&nbsp;4&nbsp;] 
                                      Main Fax Number:</td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBlack"> 
                                      <input type="text" name="MainFaxNumber" class="CastlesTextBlack" size="15" maxlength="50" value="<%=WebSiteContactInfo.Fields.Item("MainFaxNumber").Value%>">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBody">[&nbsp;5&nbsp;] 
                                      Main Email Address:</td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBlack"> 
                                      <input type="text" name="MainEmailAddress" class="CastlesTextBlack" size="25" maxlength="100" value="<%=WebSiteContactInfo.Fields.Item("MainEmailAddress").Value%>">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBody">[&nbsp;6&nbsp;] 
                                      Toll Free Phone:</td>
                                  </tr>
                                  <tr> 
                                    <td class="CastlesTextBlack"> 
                                      <input type="text" name="TollFreeTelNumber" class="CastlesTextBlack" size="15" maxlength="50" value="<%=WebSiteContactInfo.Fields.Item("TollFreeTelNumber").Value%>">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td> 
                                      <input type="hidden" name="WebSiteContactInfoID" value="<%=WebSiteContactInfo.Fields.Item("WebSiteContactInfoID").Value%>">
                                    </td>
                                  </tr>
                                </table>
                                <table width="300" border="0" cellspacing="0" cellpadding="1">
                                  <tr> 
                                    <td class="CastlesTextBlack"> 
                                      <input type="submit" name="Submit" value="Update Contact Info" class="CastlesTextBlack">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td>&nbsp;</td>
                                  </tr>
                                </table>
                              </td>
                              <td width="365" class="CastlesTextBody"> 
                                <%
If Request.QueryString("Update") = "Y" Then
%>
                                <font color="#990000">* Sucessfully Updated</font> 
                                <%
End If
%>
                              </td>
                            </tr>
                            <tr> 
                              <td width="5">&nbsp;</td>
                              <td rowspan="13" valign="top"> <br>
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
                      <tr> 
                        <td width="5">&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td width="100%" height="1"><img src="../../../../Castles/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="<%=ThinLine%>" width="100%" height="1"><img src="../../../../Castles/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="100%" height="1"><img src="../../../../Castles/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
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
