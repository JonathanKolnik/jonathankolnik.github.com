<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/templates/castlesclientcnekt.asp" -->
<%
'Displays WebSite Content
WebSiteContentID = 6
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
%>

<html><!-- #BeginTemplate "/Templates/CastlesClient.dwt" --><!-- DW6 -->
<head>

<title>Brookline, Wellesley, Weston and Newton real estate and homes for sale in Massachusetts - Castles Unlimited</title>

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
	daughter = window.open("http://216.119.112.46/manage/default.asp?Broker=Y",'daughter','toolbar=no,scrollbars=yes,resizable');
}
function registered() {
	DispWin = window.open('http://mlsplug-in.com/mlsplugin/default.asp?Office=989','b','height=420,width=650,resizable=1,scrollbars=1,menubar=0,toolbar=1');
}
//-->
</script>
<!-- #BeginEditable "script" -->
<script language="JavaScript">
<!--
function Validate(){
	{MM_Depth++};var MM_localVars = new Array();MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',74,4368)}catch(e){MM_D='\''+MM_debugError+'\''};}var errorString = ""
	;if(MM_localVars.reverseFind('errorString')==-1){MM_localVars[MM_localVars.length]='errorString';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',75,4390)}catch(e){MM_D='\''+MM_debugError+'\''};}var errorTrue = ""

	;if(MM_localVars.reverseFind('errorTrue')==-1){MM_localVars[MM_localVars.length]='errorTrue';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',77,4411)}catch(e){MM_D='\''+MM_debugError+'\''};}if (document.Subscribe.FirstName.value == "") {
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',78,4461)}catch(e){MM_D='\''+MM_debugError+'\''};}errorString=errorString + " - Please enter your first name. \r"
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',79,4527)}catch(e){MM_D='\''+MM_debugError+'\''};}errorTrue="y"
	}
	MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',81,4545)}catch(e){MM_D='\''+MM_debugError+'\''};}if (document.Subscribe.FirstName.value == "") {
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',82,4595)}catch(e){MM_D='\''+MM_debugError+'\''};}errorString=errorString + " - Please enter your last name. \r"
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',83,4660)}catch(e){MM_D='\''+MM_debugError+'\''};}errorTrue="y"
	}

	MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',86,4679)}catch(e){MM_D='\''+MM_debugError+'\''};}if (errorTrue == "y") {
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',87,4705)}catch(e){MM_D='\''+MM_debugError+'\''};}alert("The form could not be submitted due to the following: \r" + errorString) 
	}else {
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',89,4797)}catch(e){MM_D='\''+MM_debugError+'\''};}document.Subscribe.submit();
	}
MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',91,4829)}catch(e){MM_D='\''+MM_debugError+'\''};}{MM_Depth--}}

function SameAsBilling(){
	{MM_Depth++};var MM_localVars = new Array();MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',94,4859)}catch(e){MM_D='\''+MM_debugError+'\''};}var BillingAddressLine1 = document.Subscribe.BillingAddressLine1.value
	;if(MM_localVars.reverseFind('BillingAddressLine1')==-1){MM_localVars[MM_localVars.length]='BillingAddressLine1';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',95,4931)}catch(e){MM_D='\''+MM_debugError+'\''};}var BillingAddressLine2 = document.Subscribe.BillingAddressLine2.value
	;if(MM_localVars.reverseFind('BillingAddressLine2')==-1){MM_localVars[MM_localVars.length]='BillingAddressLine2';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',96,5003)}catch(e){MM_D='\''+MM_debugError+'\''};}var BillingCity = document.Subscribe.BillingCity.value
	;if(MM_localVars.reverseFind('BillingCity')==-1){MM_localVars[MM_localVars.length]='BillingCity';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',97,5059)}catch(e){MM_D='\''+MM_debugError+'\''};}var BillingStateProvinceID = document.Subscribe.BillingStateProvinceID.value
	;if(MM_localVars.reverseFind('BillingStateProvinceID')==-1){MM_localVars[MM_localVars.length]='BillingStateProvinceID';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',98,5137)}catch(e){MM_D='\''+MM_debugError+'\''};}var BillingZipPostalCode = document.Subscribe.BillingZipPostalCode.value
	;if(MM_localVars.reverseFind('BillingZipPostalCode')==-1){MM_localVars[MM_localVars.length]='BillingZipPostalCode';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',99,5211)}catch(e){MM_D='\''+MM_debugError+'\''};}var BillingCounrtyID = document.Subscribe.BillingCounrtyID.value

	;if(MM_localVars.reverseFind('BillingCounrtyID')==-1){MM_localVars[MM_localVars.length]='BillingCounrtyID';}MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',101,5278)}catch(e){MM_D='\''+MM_debugError+'\''};}if (document.Subscribe.SameAsBilling.checked = true) {
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',102,5335)}catch(e){MM_D='\''+MM_debugError+'\''};}document.Subscribe.ShippingAddressLine1.value = BillingAddressLine1
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',103,5405)}catch(e){MM_D='\''+MM_debugError+'\''};}document.Subscribe.ShippingAddressLine2.value = BillingAddressLine2
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',104,5475)}catch(e){MM_D='\''+MM_debugError+'\''};}document.Subscribe.ShippingCity.value = BillingCity
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',105,5529)}catch(e){MM_D='\''+MM_debugError+'\''};}document.Subscribe.ShippingStateProvinceID.value = BillingStateProvinceID
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',106,5605)}catch(e){MM_D='\''+MM_debugError+'\''};}document.Subscribe.ShippingZipPostalCode.value = BillingZipPostalCode
		MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',107,5677)}catch(e){MM_D='\''+MM_debugError+'\''};}document.Subscribe.ShippingCountryID.value = BillingCounrtyID
	}
MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',109,5742)}catch(e){MM_D='\''+MM_debugError+'\''};}{MM_Depth--}}

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
          <td bgcolor="#333399" height="10" width="350"><img src="../../images/clear10pixel.gif" width="10" height="10"></td>
          <td width="400" height="10"><img src="../../images/tagline.GIF" width="387" height="10"></td>
        </tr>
        <tr> 
          <td width="350"><a href="../../default.asp"><img src="../../images/cstles_logo.gif" width="151" height="91" border="0" alt="Castles Magazine"></a></td>
          <td width="400" valign="bottom" align="right"> 
            <table width="400" cellspacing="0" cellpadding="0" border="0">
              <tr> 
                <td align="right" class="CastlesTextBodyBold"><a href="javascript:OnClick=registered();" class="black">&gt; Massachusetts Homes for Sale</a></td>
                <td align="right" class="CastlesTextBodyBold"><a href="http://realestate.dreamingcode.com/weblink" class="black">&gt; Boston Condos for Sale</a></td>
                <td align="right" class="CastlesTextBodyBold"><a href="../../contactcastles" class="black">&gt; 
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
                      <td class="CastlesTextNav"><a href="../../default.asp" class="black">Home</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#A0A0A0"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="javascript:OnClick=registered();" class="black">Homes For Sale (MLS) </a></td>
                    </tr>
                    <tr>
                      <td class="CastlesTextNav"><a href="http://realestate.dreamingcode.com/weblink" class="black">Condos For Sale (LINK) </a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../FeaturedListings/default.asp" class="black">Featured 
                        Properties</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../Services/default.asp" class="black">Services We Offer</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../Reports/default.asp" class="black">Free Reports</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../Agents/default.asp" class="black">Our Agents</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../SellingTips/default.asp" class="black">Selling Tips</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../BuyingTips/default.asp" class="black">Buying Tips</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="http://www.inman.com/inmaninf/castlesunltd/index.aspx" class="black">Real Estate News</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../Franchise/default.asp" class="black">Franchise Opportunities</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../HelpFulLinks/default.asp" class="black">Helpful Links </a></td>
                    </tr>
					<!--
                    <tr> 
                      <td class="CastlesTextNav"><a href="javascript:BrokerLogin()" class="black">Broker 
                        Log-in / Place Ad</a></td>
                    </tr>
					-->
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../aboutcastles" class="black">About 
                        Castles</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="../../contactcastles" class="black">Contact 
                        Castles</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
			  <!--
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="../subscribetomagazine/" class="black">Subscribe 
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
                      <td class="CastlesTextNav"><a href="../openbrokeraccount/" class="black">Open 
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
                      <td class="CastlesTextNav"><a href="../advertisingrates/" class="black">Advertising 
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
                      <td class="CastlesTextNav"><a href="../submissionguidelines/" class="black">Submission 
                        Guidelines</a></td>
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
				-->
            </table>
          </td>
          <td width="2"><img src="../../images/clear10pixel.gif" width="2" height="1"></td>
          <td bgcolor="#CCCCCC" width="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td width="2"><img src="../../images/clear10pixel.gif" width="2" height="1"></td>
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
                          <tr> 
                            <td class="CastlesTextBody">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td class="CastlesTextBody"><%=WebSiteContentBody1%> </td>
                          </tr>
                          <tr>
                            <td class="CastlesTextBody">
                              <form name="Subscribe" method="post" action="ProcessSubscribe.asp">
                                <table width="350" border="0" cellspacing="0" cellpadding="2">
                                  <tr align="left"> 
                                    <td colspan="2" class="CastlesTextBody"><b>Subscription 
                                      Type </b></td>
                                  </tr>
                                  <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_SubscriptionType_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set SubscriptionTypes = .Execute()
End With
Set Command1 = Nothing

While (NOT SubscriptionTypes.EOF)
%>
                                  <tr align="left"> 
                                    <td class="CastlesTextBody" align="right" width="150" valign="top"> 
                                      <input type="radio" name="SubscriptionTypeID" value="<%=SubscriptionTypes.Fields.Item("SubscriptionTypeID").Value%>">
                                    </td>
                                    <td class="CastlesTextBody" width="200" valign="top"><b><%=DCFormatCurrency(SubscriptionTypes.Fields.Item("Price").Value,2)%> - <%=SubscriptionTypes.Fields.Item("SubscriptionTypeName").Value%></b><br>
                                      <%=SubscriptionTypes.Fields.Item("SubscriptionTypeShortDescription").Value%></td>
                                  </tr>
                                  <%
	SubscriptionTypes.MoveNext()
Wend
SubscriptionTypes.close
Set SubscriptionTypes = Nothing
%>
                                  <tr align="left"> 
                                    <td colspan="2" class="CastlesTextBody"><b>Your 
                                      Billing Information </b></td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">First 
                                      Name:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="FirstName" class="CastlesTextBlack" size="25" maxlength="50">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Last 
                                      Name:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="LastName" class="CastlesTextBlack" size="25" maxlength="50">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Address 
                                      Line 1:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="BillingAddressLine1" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Address 
                                      Line 2:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="BillingAddressLine2" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">City:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="BillingCity" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">State/Province:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="BillingStateProvinceID" class="CastlesTextBlack">
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
                                        <option value="<%=(StateProvinces.Fields.Item("StateProvinceID").Value)%>"><%=(StateProvinces.Fields.Item("StateProvinceName").Value)%></option>
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
                                      Code:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="BillingZipPostalCode" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Country:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="BillingCountryID" class="CastlesTextBlack">
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
%>
                                        <option value="<%=(Countries.Fields.Item("CountryID").Value)%>"><%=(Countries.Fields.Item("CountryName").Value)%></option>
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
                                      Number:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="TelNumber" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Email 
                                      Address:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="EmailAddress" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody" valign="top">Comments:</td>
                                    <td width="200" class="CastlesTextBlack" valign="top"> 
                                      <textarea name="Comments" class="CastlesTextBlack" cols="35" wrap="VIRTUAL" rows="4"></textarea>
                                    </td>
                                  </tr>
                                  <tr align="left"> 
                                    <td colspan="2" class="CastlesTextBody"><b>Your 
                                      Shipping Information</b></td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">&nbsp;</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <table width="200" border="0" cellspacing="0" cellpadding="0">
                                        <tr> 
                                          <td width="25"> 
                                            <input type="checkbox" name="SameAsBilling" value="Y" onClick="MM_bD=true;while(MM_bD&&!MM_bInEval){try{MM_bD=MM_Debug(eval(MM_D),'default.asp',424,21570)}catch(e){MM_D='\''+MM_debugError+'\''};}SameAsBilling();">
                                          </td>
                                          <td width="175" class="CastlesTextBody">Same 
                                            as Billing Information</td>
                                        </tr>
                                      </table>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Address 
                                      Line 1:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="ShippingAddressLine1" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Address 
                                      Line 2:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="ShippingAddressLine2" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">City:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="ShippingCity" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">State/Province:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="ShippingStateProvinceID" class="CastlesTextBlack">
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
                                        <option value="<%=(StateProvinces.Fields.Item("StateProvinceID").Value)%>"><%=(StateProvinces.Fields.Item("StateProvinceName").Value)%></option>
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
                                      Code:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="ShippingZipPostalCode" class="CastlesTextBlack" size="25" maxlength="100">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">Country:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="ShippingCountryID" class="CastlesTextBlack">
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
%>
                                        <option value="<%=(Countries.Fields.Item("CountryID").Value)%>"><%=(Countries.Fields.Item("CountryName").Value)%></option>
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
                                    <td colspan="2" align="left" class="CastlesTextBody"><b>Payment 
                                      Information </b></td>
                                  </tr>
                                  <tr> 
                                    <td width="150" class="CastlesTextBody" align="right">Full 
                                      Name on Card:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="CreditCardHolderName" maxlength="100" class="CastlesTextBlack" size="30">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" class="CastlesTextBody" align="right">Credit 
                                      Card Type:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="CreditCardTypeID" class="CastlesTextBlack">
                                        <option value="0">Please select a Type</option>
                                        <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_CreditCardTypes_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set CreditCardTypes = .Execute()
End With
Set Command1 = Nothing

While (NOT CreditCardTypes.EOF)
%>
                                        <option value="<%=(CreditCardTypes.Fields.Item("CreditCardTypeID").Value)%>"><%=(CreditCardTypes.Fields.Item("CreditCardTypeName").Value)%></option>
                                        <%
	CreditCardTypes.MoveNext()
Wend
CreditCardTypes.close
Set CreditCardTypes = Nothing
%>
                                      </select>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" class="CastlesTextBody" align="right">Credit 
                                      Card Number:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <input type="text" name="CreditCardNumberDisplay" maxlength="50" class="CastlesTextBlack" size="30">
                                    </td>
                                  </tr>
                                  <tr>
                                    <td width="150" class="CastlesTextBody" align="right">Credit 
                                      Card CSC Code:</td>
                                    <td width="200" class="CastlesTextBlack">
                                      <input type="text" name="CreditCardCSCCode" maxlength="50" class="CastlesTextBlack" size="10">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" class="CastlesTextBody" align="right">Expiration 
                                      Date:</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <select name="CreditCardExpirationMonth" class="CastlesTextBlack">
                                        <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_Months_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Months = .Execute()
End With
Set Command1 = Nothing

While (NOT Months.EOF)
%>
                                        <option value="<%=(Months.Fields.Item("MonthValue").Value)%>"><%=(Months.Fields.Item("MonthValue").Value)%></option>
                                        <%
	Months.MoveNext()
Wend
Months.close
Set Months = Nothing
%>
                                      </select>
                                      &nbsp;&nbsp; 
                                      <select name="CreditCardExpirationYear" class="CastlesTextBlack">
                                        <%
Set Command1 = Server.CreateObject("ADODB.Command")
With Command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_Years_List"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set Years = .Execute()
End With
Set Command1 = Nothing

While (NOT Years.EOF)
%>
                                        <option value="<%=(Years.Fields.Item("YearValue").Value)%>"><%=(Years.Fields.Item("YearValue").Value)%></option>
                                        <%
	Years.MoveNext()
Wend
Years.close
Set Years = Nothing
%>
                                      </select>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">&nbsp;</td>
                                    <td width="200" class="CastlesTextBlack">&nbsp;</td>
                                  </tr>
                                  <tr> 
                                    <td width="150" align="right" class="CastlesTextBody">&nbsp;</td>
                                    <td width="200" class="CastlesTextBlack"> 
                                      <table width="40" border="0" cellspacing="0" cellpadding="3">
                                        <tr> 
                                          <td class="CastlesTextBodyBold" align="center" bgcolor="#D6D4BA">&gt;<a href="javascript:onClick=Validate();" class="normal">Subscribe</a></td>
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
          <td width="150" valign="top" bgcolor="#FFFFFF" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#FFFFFF" width="196" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="../../images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td bgcolor="#FFFFFF" width="397" height="2"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
        </tr>
        <tr> 
          <td bgcolor="#CCCCCC" width="150" valign="top" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="196" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="396" height="1"><img src="../../images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td width="150" valign="top" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="196" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
          <td width="396" height="2" bgcolor="#FFFFFF"><img src="../../images/clear10pixel.gif" width="1" height="2"></td>
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
          <td width="595" class="CastlesTextBody">&copy; 2003 <a href="http://www.castlesunltd.com" class="normal">Castles 
            Unlimited</a> &nbsp;&nbsp;All rights reserved. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="CastlesTextBodyBold">&gt;<a href="../../misc/privacypolicy.asp" class="normal">Privacy 
            Policy</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;<a href="../../misc/termsofuse.asp" class="normal">Terms 
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
