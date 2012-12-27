<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 0
PageTopNavigationSubHeaderID = 0
EntityID = 0
OriginatingEntityID = Request.QueryString("OriginatingEntityID")
EntityPrimaryKeyValue = Request.QueryString("EntityPrimaryKeyValue")
%>
<!--#include virtual="/templates/castlessystemcnektonly.asp" -->
<%

'---------------Begin Page Level Multilingual Translation-----------------------
'WordPhrasesOnPage = "Actions(|)Active(|)Add an Address Book Entry(|)Add Entry(|)Address Book Entry Name(|)Address Line 2(|)City(|)Country(|)Delete(|)Edit(|)Hide Help Text(|)Location Type(|)No Results Found(|)Show Help Text(|)Special Instructions(|)State/Province(|)Street Name(|)Street Number(|)Zip/Postal Code"

'WordPhrasesOnPageArray = Split(WordPhrasesOnPage,"(|)")
'WhereClause = ""
'WordPhraseCount = 1

'For Each WordPhrase In WordPhrasesOnPageArray
'    If WordPhraseCount = 1 Then
'        WhereClause = WhereClause & " EnglishTranslation = '" & WordPhrase & "' AND LanguageID = " & LanguageID 
'    Else
'        WhereClause = WhereClause & " OR EnglishTranslation = '" & WordPhrase & "' AND LanguageID = " & LanguageID 
'    End If
'    WordPhraseCount = WordPhraseCount + 1
'Next

'Set Command1 = Server.CreateObject("ADODB.Command")
'With Command1
'    .ActiveConnection = Connect
'    .CommandText = "Castles_TranslateWordPhrases_For_Page"
'    .Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
'    .Parameters.Append .CreateParameter("@WhereClause", 201, 1,20000,WhereClause)
'    .CommandType = 4
'    .CommandTimeout = 0
'    .Prepared = true
'    Set TranslationResults = .Execute()
'End With
'Set Command1 = Nothing

'TranslationResultsArray = TranslationResults.getrows
'TranslationResults.close
'Set TranslationResults = Nothing
'TranslationResultsArrayNumRows = uBound(TranslationResultsArray,2)
'TranslateCount = 0
'Field_TranslatedWordPhrase = 0

'WordPhrase_Actions = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_Active = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_AddAddressBookEntry = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_AddEntry = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_AddressBookEntryName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_AddressLine2 = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_City = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_Country = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_Delete = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_Edit = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_HideHelpText = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_LocationType = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_NoResultsFound = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_ShowHelpText = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_SpecialInstructions = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_StateProvince = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_StreetName = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_StreetNumber = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'TranslateCount = TranslateCount + 1
'WordPhrase_ZipPostalCode = TranslationResultsArray(Field_TranslatedWordPhrase,TranslateCount)
'---------------End Multilingual Translation----------------------- 


%>
<html>
<head>
<title>Castles - Management System</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.CastlesTextBlack {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 7pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBlackBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBody {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 7pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextBody%>; text-decoration: none}
.CastlesTextBodyBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextBody%>; text-decoration: none}
.CastlesTextHeader {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 7pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextHeader%>; text-decoration: none}
.CastlesTextHeaderBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextHeader%>; text-decoration: none}
.CastlesTextWhite {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 7pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextWhiteBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextNavDark {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 7pt; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: <%=TextNavHighlight%>; text-decoration: none}

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
function Validate(){
	var errorString = ""
	var errorTrue = ""

	if (document.ForgotPassword.EmailAddress.value == "") {
		errorString=errorString + " - Please enter your email address. \r"
		errorTrue="y"
	}
	if (errorTrue == "y") {
		alert("Missing Required Fields: \r" + errorString) 
		return false;
	}else {
		return true;
	}
}

function DeleteCheckedRecords(){
	if(confirm("Are you sure you want to permanently delete the selected item(s)?")){
		return true;
	}else {
		return false;
	}
}

//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="<%=TopBar%>"> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="35%"><img src="images/Castles_logo.gif" width="190" height="40" usemap="#Map" border="0"></td>
                <td width="65%">&nbsp;</td>
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
          <td bgcolor="<%=NavDarkBorder%>" height="1"><img src="administration/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="administration/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavRegular%>" height="19"><img src="administration/images/clear10pixel.gif" width="1" height="19"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavDarkBorder%>" height="1"><img src="administration/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="administration/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="19"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td class="CastlesTextNavDark">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="<%=NavMuted%>" height="1"><img src="administration/images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td class="CastlesTextBody" colspan="3" height="8"><img src="images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td class="CastlesTextBody" width="3">&nbsp;</td>
                <td class="CastlesTextBodyBig" width="80%"><b>Forgot Password</b><br><br>Please enter 
                  the Email Address you used when you signed up. Your password 
                  will be sent to that Email Address.</td>
				<td class="CastlesTextBody" width="10%" align="right">&nbsp;</td>
              </tr>
              <%
If SystemHelpContentText <> "" Then
%>
              <tr> 
                <td class="CastlesTextBody" width="3">&nbsp;</td>
                <td class="CastlesTextBody" colspan="2">&nbsp;<%=SystemHelpContentText%></td>
              </tr>
              <%
End If
%>
              <tr> 
                <td class="CastlesTextBody" colspan="3" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="<%=TitleBar%>"><img src="images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td> 
		  <%
		  if request.QueryString("Success") <> "Y" then
		  %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <form name="ForgotPassword" onSubmit="return Validate();" method="post" action="driveforgotpassword.asp">
                <tr> 
                  <td colspan="5" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="20%" class="CastlesTextBody">&nbsp;</td>
                  <td class="CastlesTextBody" colspan="2"><font color="red">
				  <%
				  if request.querystring("success") = "N" then
				  	response.write "Your email address was not found."
				end if
				%>
				  </font></td>
                  <td width="30%" class="CastlesTextBody">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="20%" class="CastlesTextBody">&nbsp;</td>
                  <td class="CastlesTextBody" colspan="2">[&nbsp;1&nbsp;]&nbsp;Email 
                    Address:</td>
                  <td width="30%" class="CastlesTextBody">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="20%" class="CastlesTextBody">&nbsp;</td>
                  <td class="CastlesTextBody" colspan="2"> 
                    <input type="text" name="EmailAddress" size="40" maxlength="100" class="CastlesTextBlack">
                  </td>
                  <td width="30%" class="CastlesTextBody"> 
                    <input type="submit" value="Submit" name="submit"  class="CastlesTextBody">
                  </td>
                </tr>
                <tr> 
                  <td align="right" width="20%" class="CastlesTextBody">&nbsp;</td>
                  <td class="CastlesTextBody" width="30%">&nbsp;</td>
                  <td class="CastlesTextBody" width="20%">&nbsp;</td>
                  <td class="CastlesTextBody" width="30%">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="5" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
              </form>
            </table>
			<%
			else
			%>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td colspan="5" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="20%" class="CastlesTextBody">&nbsp;</td>
                  <td class="CastlesTextBody" colspan="2">&nbsp;</td>
                  <td width="30%" class="CastlesTextBody">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="20%" class="CastlesTextBody">&nbsp;</td>
                  <td class="CastlesTextBody" colspan="2"><B>Your Account was successfully located and your password has been sent to your email address.</b></td>
                  <td width="30%" class="CastlesTextBody">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="20%" class="CastlesTextBody">&nbsp;</td>
                  
                <td class="CastlesTextBody" colspan="2"><a href="javascript: window.close()" class="normal">Close this Window</a></td>
                  <td width="30%" class="CastlesTextBody">&nbsp; 
                  </td>
                </tr>
                <tr> 
                  <td align="right" width="20%" class="CastlesTextBody">&nbsp;</td>
                  <td class="CastlesTextBody" width="30%">&nbsp;</td>
                  <td class="CastlesTextBody" width="20%">&nbsp;</td>
                  <td class="CastlesTextBody" width="30%">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="5" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
            </table>
			<%
			end if
			%>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      </td>
  </tr>
  <tr width="100%"> 
    <td bgcolor="<%=TitleBar%>" width="100%" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
  </tr>
  <tr width="100%"> 
    <td width="100%" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td > 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="25%">&nbsp;</td>
                <td width="75%" class="CastlesTextBody">&copy; 2003 . 
                  All rights reserved.</td>
              </tr>
            </table>
          </td>
          <td align="right"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="75%" align="right"><a href="http://www.dreamingcode.com" class="normal"><img src="images/dc_logo_footer.jpg" width="97" height="23" alt="DreamingCode, Inc." border="0"></a></td>
                <td width="25%" class="CastlesTextBody">&nbsp;</td>
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
