<%@LANGUAGE="VBSCRIPT"%>
<%
PageTopNavigationHeaderID = 2
PageTopNavigationSubHeaderID = 7
FromPageTopNavigationSubHeaderID = 6
EntityID = 2
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
WebSiteContentName = Request.QueryString("WebSiteContentName")

'WebSite Content Info
WebSiteContentID = Request.QueryString("WebSiteContentID")
WebSiteContentLanguageID = Request.QueryString("WebSiteContentLanguageID")
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_System_WebSiteContent_Display_For_Edit"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@WebSiteContentID", 200, 1,200,WebSiteContentID)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,WebSiteContentLanguageID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	Set WebSiteContentProfile = .Execute()
End With
Set Command1 = Nothing

ActiveLanguageList = ActiveLanguages()
LanguageListArray = Split(ActiveLanguageList,",")
LanguagesCount = 1

For Each Language In LanguageListArray
	LanguageSpecificArray = Split(Language,"<!DCDELIMETER!>")
	SpecificLanguageName = LanguageSpecificArray(0)
	SpecificLanguageID = LanguageSpecificArray(1)
	If LanguagesCount <> 1 Then
		LanguageList = LanguageList & "&nbsp;&nbsp;-&nbsp;&nbsp;<a href=""editwebsitecontent.asp?WebSiteContentID=" & WebSiteContentID & "&WebSiteContentLanguageID=" & SpecificLanguageID & """ class=""normal"">" & SpecificLanguageName & "</a>"
	Else
		LanguageList = LanguageList & "<a href=""editwebsitecontent.asp?WebSiteContentID=" & WebSiteContentID & "&WebSiteContentLanguageID=" & SpecificLanguageID & """ class=""normal"">" & SpecificLanguageName & "</a>"
	End If
	LanguagesCount = LanguagesCount + 1
Next
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

	if (document.EditEntity.WebSiteContentCaption1.value.length > 2000) {
		errorString=errorString + "You are only allowed 2000 characters for this field. \r"
		errorTrue="y"
	}
	if (errorTrue == "y") {
		alert(errorString) 
		return false;
	}else {
		return true;
	}
}
//-->
</script>
<script language="Javascript1.2">
<!-- 
_editor_url = "http://www.dreamingcode.com/htmlarea/";   
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);

if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
	if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
		if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
			if (win_ie_ver >= 5.5) {
				 document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
				 document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { 
	document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); 
}
// -->
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
                <td bgcolor="<%=TitleBar%>" width="100%" class="CastlesTextHeader" height="20">&nbsp;&nbsp;&nbsp;<b><%=UCase(WebSiteContentName & " " & WordPhrase_WebSiteContent)%></b></td>
              </tr>
              <form name="EditEntity" method="post" onSubmit="return Validate();" action="drivewebsitecontent.asp?DCDataDriverType=SQLUpdate">
                <tr> 
                  <td width="100%" height="1"><img src="../../../../Castles/manage/administration/images/clear10pixel.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="100%" bgcolor="<%=LightField%>"> 
                    <table width="750" border="0" cellspacing="0" cellpadding="2">
                      <tr> 
                        <td width="5">&nbsp;</td>
                        <td rowspan="14" valign="top"> 
                          <table width="600" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td class="CastlesTextBody"><b><%=WordPhrase_Languages%></b>:&nbsp;&nbsp;<%=LanguageList%></td>
                            </tr>
<%
CaptionCount = 1
For i = 1 to 10
	WebSiteContentCaption = WebSiteContentProfile.Fields.Item("WebSiteContentCaption" & CaptionCount).Value
	If WebSiteContentCaption <> "<!DCNOCONTENT!>" OR isNull(WebSiteContentCaption) Then
%>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;<%=CaptionCount%>&nbsp;A&nbsp;]&nbsp;<%=WordPhrase_WebSiteContentCaptionHeader%>&nbsp;<%=CaptionCount%></td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="text" name="WebSiteContentCaptionHeader<%=CaptionCount%>" class="CastlesTextBlack" size="40" value="<%=WebSiteContentProfile.Fields.Item("WebSiteContentCaptionHeader" & CaptionCount).Value%>" maxlength="30">
                                &nbsp;&nbsp; </td>
                            </tr>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;<%=CaptionCount%>&nbsp;B&nbsp;]&nbsp;<%=WordPhrase_WebSiteContentCaption%>&nbsp;<%=CaptionCount%></td>
                            </tr>
                            <script language="JavaScript1.2" defer>
									var config = new Object(); // create new config object
									config.width = "90%";
									config.height = "200px";
									config.bodyStyle = 'background-color: white; font-family: "Verdana"; font-size: x-small;';
									config.debug = 0;
									
									config.toolbar = [
									  //['fontname'],
									 // ['fontsize'],
									 // ['fontstyle'],
									//  ['linebreak'],
									  ['bold','italic','underline','separator'],
									  ['strikethrough','subscript','superscript','separator'],
									  ['justifyleft','justifycenter','justifyright','separator'],
									  ['OrderedList','UnOrderedList','separator'],
									 // ['forecolor','backcolor','separator'],
									//['custom1','custom2','custom3','separator'],
									  ['HorizontalRule','Createlink','InsertImage','htmlmode','separator'],
									  ['help']
									]; 


									editor_generate('WebSiteContentCaption<%=CaptionCount%>',config);
								</script>
							<tr> 
                              <td class="CastlesTextBlack"> 
                                <textarea name="WebSiteContentCaption<%=CaptionCount%>" class="CastlesTextBlack" cols="70" wrap="VIRTUAL" rows="7"><%=WebSiteContentProfile.Fields.Item("WebSiteContentCaption" & CaptionCount).Value%></textarea>
                                &nbsp;&nbsp; </td>
                            </tr>
<%
	End If
	CaptionCount = CaptionCount + 1
Next
%>
                            <tr> 
                              <td class="CastlesTextBody">&nbsp;</td>
                            </tr>
<%
BodyCount = 1
For i = 1 to 10
	CombinedCount = CaptionCount + BodyCount
	WebSiteContentBody = WebSiteContentProfile.Fields.Item("WebSiteContentBody" & BodyCount).Value
	If WebSiteContentBody <> "<!DCNOCONTENT!>" OR isNull(WebSiteContentBody) Then
%>
                            <tr> 
                              <td class="CastlesTextBody">[&nbsp;<%=CombinedCount%>&nbsp;]&nbsp;<%=WordPhrase_WebSiteContentBody%>&nbsp;<%=BodyCount%></td>
                            </tr>
                            <script language="JavaScript1.2" defer>
									var config = new Object(); // create new config object
									config.width = "90%";
									config.height = "200px";
									config.bodyStyle = 'background-color: white; font-family: "Verdana"; font-size: x-small;';
									config.debug = 0;
									
									config.toolbar = [
									  //['fontname'],
									 // ['fontsize'],
									 // ['fontstyle'],
									//  ['linebreak'],
									  ['bold','italic','underline','separator'],
									  ['strikethrough','subscript','superscript','separator'],
									  ['justifyleft','justifycenter','justifyright','separator'],
									  ['OrderedList','UnOrderedList','separator'],
									 // ['forecolor','backcolor','separator'],
									//['custom1','custom2','custom3','separator'],
									  ['HorizontalRule','Createlink','InsertImage','htmlmode','separator'],
									  ['help']
									]; 


									editor_generate('WebSiteContentBody<%=BodyCount%>',config);
								</script>
							<tr> 
                              <td class="CastlesTextBlack"> 
                                <textarea name="WebSiteContentBody<%=BodyCount%>" class="CastlesTextBlack" cols="70" wrap="VIRTUAL" rows="10"><%=WebSiteContentProfile.Fields.Item("WebSiteContentBody" & BodyCount).Value%></textarea>
                                &nbsp;&nbsp; </td>
                            </tr>
<%
	End If
	BodyCount = BodyCount + 1
Next
%>
                            <tr> 
                              <td class="CastlesTextBody">&nbsp;</td>
                            </tr>
                          </table>
                          <br>
                          <table width="300" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td class="CastlesTextBlack"> 
                                <input type="hidden" name="UniqueID" value="<%=WebSiteContentProfile.Fields.Item("UniqueID").Value%>">
                                <input type="submit" name="Submit" value="<%=WordPhrase_EditWebSiteContent%>" class="CastlesTextBlack">
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
