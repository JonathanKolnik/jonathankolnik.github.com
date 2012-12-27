<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/templates/castlesclientcnekt.asp" -->
<%
'Displays WebSite Content
WebSiteContentID = 7
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_ClientSide_WebSiteContent_Detail"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4) 
	.Parameters.Append .CreateParameter("@WebSiteContentID", 200, 1,200,WebSiteContentID)
	.Parameters.Append .CreateParameter("@LanguageID", 200, 1,200,LanguageID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set WebSiteContent = .Execute()
End With
Set Command1 = Nothing

If Not WebSiteContent.EOF Then
	WebSiteContentName = WebSiteContent.Fields.Item("WebSiteContentName").Value
	WebSiteContentCaption1 = WebSiteContent.Fields.Item("WebSiteContentCaption1").Value
	WebSiteContentCaptionHeader1 = WebSiteContent.Fields.Item("WebSiteContentCaptionHeader1").Value

	If Len(WebSiteContentCaption1) <> 0 Then
		WebSiteContentCaption1 = Replace(WebSiteContentCaption1,vbcrlf,"<br>")
	End If
	
	WebSiteContentBody1 = WebSiteContent.Fields.Item("WebSiteContentBody1").Value
	If Len(WebSiteContentBody1) <> 0 Then
		WebSiteContentBody1 = Replace(WebSiteContentBody1,vbcrlf,"<br>")
	End If
End If

dupusername = request("dupusername")
CompanyName = request("CompanyName")
FirstName = request("FirstName")
MiddleInitial = request("MiddleInitial")
LastName = request("LastName")
AddressLine1 = request("AddressLine1")
AddressLine2 = request("AddressLine2")
City = request("City")
StateProvinceID = request("StateProvinceID")
ZipPostalCode = request("ZipPostalCode")
if request("CountryID") = "" then
	CountryID = 1
else
	CountryID = request("CountryID")
end if
TelNumber = request("TelNumber")
FaxNumber = request("FaxNumber")
EmailAddress = request("EmailAddress")
UserName = request("UserName")
if dupusername = "Y" then
	msg = "** The username <font color=red size=2>" & UserName & "</font> already exist. Please select another one."
end if
%>

<html><!-- #BeginTemplate "/Templates/CastlesClient.dwt" --><!-- DW6 -->
<head>

<title>Castles Magazine - The International Magazine for Distinctive Properties</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.CastlesTextBlack {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 400; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBlackBold {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextWhite {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 400; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextWhiteHeader {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextWhiteBold {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextNav {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #333333; text-decoration: none}
.CastlesTextBody {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 400; font-variant: normal; text-transform: none; color: #666633; text-decoration: none}
.CastlesTextBodyBold {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: 800; font-variant: normal; text-transform: none; color: #666633; text-decoration: none}

A.normal:link    { text-decoration: none; color: "#666633"; font-weight: 800}
A.normal:visited { text-decoration: none; color: "#666633"; font-weight: 800}
A.normal:active  { text-decoration: none; color: "#666633"; font-weight: 800}
A.normal:hover   { text-decoration: underline; color: "#666633"; font-weight: 800}

A.white:link    { text-decoration: none; color: "#FFFFFF"; font-weight: 800}
A.white:visited { text-decoration: none; color: "#FFFFFF"; font-weight: 800}
A.white:active  { text-decoration: none; color: "#FFFFFF"; font-weight: 800}
A.white:hover   { text-decoration: underline; color: "#FFFFFF"; font-weight: 800}

A.black:link    { text-decoration: none; color: "#333333"; font-weight: 800}
A.black:visited { text-decoration: none; color: "#333333"; font-weight: 800}
A.black:active  { text-decoration: none; color: "#333333"; font-weight: 800}
A.black:hover   { text-decoration: underline; color: "#333333"; font-weight: 800}
-->
</style>
<script language="JavaScript">
<!--
function BrokerLogin(){
	daughter = window.open("http://www.castlesmag.com/manage/default.asp?Broker=Y",'daughter','toolbar=Yes,width=600,height=400,top=50,left=50,scrollbars=yes,resizable');
}
//-->
</script>
<!-- #BeginEditable "script" -->
<script language="JavaScript">
<!--
function IsAlphaNumeric(field){
	var ValidChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_@.";
	
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

// Email validation
function ValidateEmailAddress(Oform,strTest,Ofield1,Ofield2)
	
	{

		strValidating = "";

		strError=""
		strFieldname=Ofield1.name;	//the name of the Ofield1
		strFieldval=Ofield1.value;	//the value of the Ofield1

		
		if(strValidating && strValidating != strFieldname){return false;}	//if we are strValidating a specific field, do not process calls coming from another field's onBlur
		if(strTest=="email"){strError+=validateEmail(Ofield1) + "\n";}
		
			if(strError.length>5)
		{
			strValidating = strFieldname;
			alert(strError)
			var Oformname=Oform.name
			var evalstring="document." + Oformname + "." + strFieldname;
			//eval(evalstring + ".value=strNewVal;")
			eval(evalstring + ".focus();")
			eval(evalstring + ".select();")
		}
		else
		{
			strValidating = false;
		}
	}

function validateEmail(arg) 
	{		
		if(strFieldval == "")
		{
		 return ""
		}
				
		var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
		var regex = new RegExp(emailReg);

		var bResult = regex.test(strFieldval);
		
		if(bResult == false)
		{
			strNewVal="";
			return "This field requires a email with valid format.\nex: bgates@msn.net";
		}
		return ""
	}

// end of Email validation
	
function Validate(){
	var errorString = ""
	var errorTrue = ""
	if (IsAlphaNumeric(document.OpenBrokerAccount.UserName.value) == false) {
		errorString=errorString + " - UserName may only contain AlphaNumeric and _.@ characters only. \r";
		errorTrue="y";
	}

	if (document.OpenBrokerAccount.FirstName.value == "") {
		errorString=errorString + " - Please enter your first name. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.AddressLine1.value == "") {
		errorString=errorString + " - Please enter your  address line 1. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.City.value == "") {
		errorString=errorString + " - Please enter your  city. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.StateProvinceID.value == "") {
		errorString=errorString + " - Please select state. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.ZipPostalCode.value == "") {
		errorString=errorString + " - Please enter your  zip/postal code. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.TelNumber.value == "") {
		errorString=errorString + " - Please enter your phone number. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.EmailAddress.value == "") {
		errorString=errorString + " - Please enter your email address. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.UserName.value == "") {
		errorString=errorString + " - Please enter UserName. \r"
		errorTrue="y"
	}
	if (document.OpenBrokerAccount.Password.value == "") {
		errorString=errorString + " - Please enter Password. \r"
		errorTrue="y"
	}
	if (errorTrue == "y") {
		alert("The form could not be submitted due to the following: \r" + errorString) 
	}else {
		document.OpenBrokerAccount.submit();
	}
}


//-->
</script>
<!-- #EndEditable -->

</head>
<body bgcolor="#FFFFFF" leftmargin="3" topmargin="3" marginwidth="3" marginheight="3">
<table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#CE6F19" height="10" width="350"><img src="../images/clear10pixel.gif" width="10" height="10"></td>
          <td width="400" height="10"><img src="../images/tagline.GIF" width="387" height="10"></td>
        </tr>
        <tr> 
          <td width="350"><a href="/default.asp"><img src="../images/cstles_logo.gif" width="338" height="60" border="0" alt="Castles Magazine"></a></td>
          <td width="400" valign="bottom" align="right"> 
            <table width="400" cellspacing="0" cellpadding="0" border="0">
              <tr> 
                <td width="300" align="right" class="CastlesTextBodyBold"><a href="javascript:BrokerLogin()" class="black">&gt; 
                  Click here to place a new listing (Brokers)</a></td>
                <td width="100" align="right" class="CastlesTextBodyBold"><a href="../contactcastles/default.asp" class="black">&gt; 
                  Contact Us</a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="150" valign="top" bgcolor="#CCCCCC"> 
            <table width="150" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="/default.asp" class="black">Home</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#A0A0A0"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="/Listings/default.asp" class="black">Search 
                        Properties</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="/FeaturedListings/default.asp" class="black">Featured 
                        Properties</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="javascript:BrokerLogin()" class="black">Broker 
                        Log-in / Place Ad</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../aboutcastles" class="black">About 
                        Castles</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../contactcastles" class="black">Contact 
                        Castles</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../subscribetomagazine" class="black">Subscribe 
                        to Magazine</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../openbrokeraccount" class="black">Open 
                        a Broker Account</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../advertisingrates" class="black">Advertising 
                        Rates</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>

              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../submissionguidelines" class="black">Submission 
                        Guidelines</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
			 <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../images/castles_mls.pdf" target="_blank" class="black"><font color="red">Brokers: Power your Website</Font></a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>			
              <tr bgcolor="#CCCCCC"> 
                <td width="10">&nbsp;</td>
                <td width="130" height="50">&nbsp;</td>
                <td width="10">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
          <td bgcolor="#CCCCCC" width="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
          <td width="595" valign="top" colspan="4"><!-- #BeginEditable "body" -->
            <table width="595" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="395" valign="top"> 
                  <table width="395" border="0" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#CC6600"> 
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                      <td width="375"> 
                        <table width="355" border="0" cellspacing="0" cellpadding="3" height="20">
                          <tr> 
                            <td class="CastlesTextWhiteBold"><%=WebSiteContentName%></td>
                          </tr>
                        </table>
                      </td>
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                    </tr>
                    <tr> 
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                      <td width="375" valign="top"> 
                        <table width="375" border="0" cellspacing="0" cellpadding="3">
                          <%if dupusername="Y" then%>
							  <tr> 
								<td class="CastlesTextBody">&nbsp;</td>
							  </tr>
							  <tr> 
								<td class="CastlesTextBody"><%=msg%> </td>
							  </tr>
						  <%end if%>
						  <tr> 
                            <td class="CastlesTextBody">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td class="CastlesTextBody"><%=WebSiteContentBody1%> </td>
                          </tr>
                          <tr>
                            <td class="CastlesTextBody">
                              <form name="OpenBrokerAccount" method="post" action="drivebrokeraccounts.asp?DCDataDriverType=SQLInsert">
                                <table width="350" border="0" cellspacing="0" cellpadding="2">
                                  <tr align="left"> 
                                    <td colspan="2" class="CastlesTextBody"><b>Your 
                                      Information </b></td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Company 
                                      Name: </td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="CompanyName" value="<%=CompanyName%>" class="CastlesTextBlack" size="30" maxlength="200">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">First 
                                      Name*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="FirstName" value="<%=FirstName%>" class="CastlesTextBlack" size="25" maxlength="50">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Middle 
                                      Initial: </td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="MiddleInitial" value="<%=MiddleInitial%>" class="CastlesTextBlack" size="5" maxlength="1">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Last 
                                      Name:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="LastName" value="<%=LastName%>" class="CastlesTextBlack" size="25" maxlength="50">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Address 
                                      Line 1*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="AddressLine1" value="<%=AddressLine1%>" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Address 
                                      Line 2:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="AddressLine2" value="<%=AddressLine2%>" class="CastlesTextBlack" size="20" maxlength="50">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">City*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="City" value="<%=City%>" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">State*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="StateProvinceID" class="CastlesTextBlack">
                                        <option value="">Please Select a State</option>
                                        <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_StateProvince_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set StateProvinces = .Execute()
End With
Set Command1 = Nothing

While (NOT StateProvinces.EOF)
%>
                                        <option value="<%=(StateProvinces.Fields.Item("StateProvinceID").Value)%>" <%if trim(StateProvinces.Fields.Item("StateProvinceID").Value) = trim(StateProvinceID) then%>Selected<%end if%>><%=(StateProvinces.Fields.Item("StateProvinceName").Value)%></option>
                                        <%
	StateProvinces.MoveNext()
Wend
StateProvinces.close
Set StateProvinces = Nothing
%>
                                      </select>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Zip/Postal 
                                      Code*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="ZipPostalCode" value="<%=ZipPostalCode%>" class="CastlesTextBlack" size="15" maxlength="50">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Country:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="CountryID" class="CastlesTextBlack">
                                        <option value="0">Please Select a Country</option>
                                        <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_Country_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Countries = .Execute()
End With
Set Command1 = Nothing

While (NOT Countries.EOF)
	if trim(Countries.Fields.Item("CountryID").Value) = trim(CountryID) then
		makeSelected = "selected"
	else
		makeSelected = ""
	end if
%>
                                        <option value="<%=(Countries.Fields.Item("CountryID").Value)%>" <%=makeSelected%>><%=(Countries.Fields.Item("CountryName").Value)%></option>
                                        <%
	Countries.MoveNext()
Wend
Countries.close
Set Countries = Nothing
%>
                                      </select>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Phone 
                                      Number*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="TelNumber" value="<%=TelNumber%>" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Fax 
                                      Number:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="FaxNumber" value="<%=FaxNumber%>" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Email 
                                      Address*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="EmailAddress" value="<%=EmailAddress%>" class="CastlesTextBlack" onblur="ValidateEmailAddress(this.form,'email',this);" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Username*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="UserName" class="CastlesTextBlack" size="20" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Password*:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="password" name="Password" class="CastlesTextBlack" size="15" maxlength="10">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">&nbsp;</td>
                                    <td width="200" class="CastlesTextBlack">&nbsp;</td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">&nbsp;</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <table width="90" border="0" cellspacing="0" cellpadding="3">
                                        <tr> 
                                          <td class="CastlesTextBodyBold" align="center" bgcolor="#D6D4BA">&gt;<a href="javascript:onClick=Validate();" class="normal">Open 
                                            Account </a></td>
                                        </tr>
                                      </table>
                                    </td>
                                  </tr>
                                </table>
                  </form>
								</td>
                          </tr>
                          <tr> 
                            <td class="CastlesTextBody">&nbsp;</td>
                          </tr>
                        </table>
                      </td>
                      <td width="10"><img src="../images/clear10pixel.gif" width="10" height="1"></td>
                    </tr>
                  </table>
                </td>
                <td bgcolor="#FFFFFF" width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
                <td bgcolor="#CCCCCC" width="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                <td bgcolor="#FFFFFF" width="2"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
                <td width="195" valign="top" bgcolor="#EDEBDB" height="100%"> 
                  <table width="195" border="0" cellspacing="0" cellpadding="0">
                    <form name="PropertyQuickSearch" method="post" action="">
                      <tr> 
                        <td width="10" bgcolor="#D6D4BA">&nbsp;</td>
                        <td width="176" bgcolor="#D6D4BA"> 
                          <table width="170" border="0" cellspacing="0" cellpadding="3" height="20">
                            <tr> 
                              <td class="CastlesTextBodyBold"><%=WebSiteContentCaptionHeader1%></td>
                            </tr>
                          </table>
                        </td>
                        <td width="10" bgcolor="#D6D4BA">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="176" class="CastlesTextBody">&nbsp;</td>
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="176" class="CastlesTextBlack">
                          <table width="175" border="0" cellspacing="0" cellpadding="3" height="20">
                            <tr> 
                              <td class="CastlesTextBody"><%=WebSiteContentCaption1%></td>
                            </tr>
                          </table>
                        </td>
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="176" class="CastlesTextBodyBold">&nbsp;</td>
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                      </tr>
                    </form>
                  </table>
                </td>
              </tr>
            </table>
            <!-- #EndEditable --></td>
        </tr>
        <tr> 
          <td width="150" valign="top" bgcolor="#FFFFFF" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#FFFFFF" width="196" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td bgcolor="#FFFFFF" width="397" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
        </tr>
        <tr> 
          <td bgcolor="#CCCCCC" width="150" valign="top" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="196" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="396" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td width="150" valign="top" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="196" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="396" height="2" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
  	<td width="750">
		<table width="750" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="150">&nbsp;</td>
				<td width="5">&nbsp;</td>
				
          <td width="595"><!-- #BeginEditable "subBody" --><!-- #EndEditable --></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr> 
    <td valign="top"> 
      <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="154">&nbsp;</td>
          <td width="595" class="CastlesTextBody">&copy; <%=year(Date())%> <a href="http://www.castlesmag.com" class="normal">Castles 
            Magazine</a> &nbsp;&nbsp;All rights reserved. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="CastlesTextBodyBold">&gt;<a href="../misc/privacypolicy.asp" class="normal">Privacy 
            Policy</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;<a href="../misc/termsofuse.asp" class="normal">Terms 
            of Use</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;
            </span></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
</body>
<!-- #EndTemplate --></html>
