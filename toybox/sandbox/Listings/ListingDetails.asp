<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/templates/castlesclientcnekt.asp" -->
<%
Function findValue(src)
	dest = request.Form(src)
	if Len(dest)=0 then
		dest = request.QueryString(src)
	end if
	findValue = dest
End Function

SearchStateProvinceID = findValue("SearchStateProvinceID")
SearchCity = findValue("SearchCity")
SearchZipcode = findValue("SearchZipcode")
SearchPriceFrom = findValue("SearchPriceFrom")
SearchPriceTo = findValue("SearchPriceTo")
SearchWaterfront = findValue("SearchWaterfront")
SearchSki = findValue("SearchSki")
SearchCondo = findValue("SearchCondo")
SearchResort = findValue("SearchResort")
SearchCountryClub = findValue("SearchCountryClub")
SearchFarmOrRanch = findValue("SearchFarmOrRanch")
Featured = findValue("Featured")
ContactSuccess = Request.QueryString("ContactSuccess")
ListingID = Request.QueryString("ListingID")
ZipListing = findValue("ZipListing")

Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_ClientSide_Listing_Details"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4) 
	.Parameters.Append .CreateParameter("@ListingID", 200, 1,200,ListingID)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set ListingProfile = .Execute()
End With
Set Command1 = Nothing

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
		.CommandText = "Castles_ClientSide_DisplayFeatureNames"
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
	if((fieldvalue<>"")and(fieldvalue<>"N/A")and(fieldvalue<>"0")and(fieldvalue<>"N")and(not isNull(fieldvalue))) then
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
function Validate(){
	var UserName = document.ContactBroker.UserName.value;
	var TelNumber = document.ContactBroker.TelNumber.value;
	var EmailAddress = document.ContactBroker.EmailAddress.value;
	if((UserName=="")||(TelNumber=="")||(EmailAddress=="")){
		alert("Your Name, Tel. Number and Email Address are required.");
		return false;
	}else{
		document.ContactBroker.action="driveContactBroker.asp?SearchStateProvinceID=<%=Server.URLEncode(SearchStateProvinceID)%>&SearchCity=<%=Server.URLEncode(SearchCity)%>&SearchZipcode=<%=Server.URLEncode(SearchZipcode)%>&SearchPriceFrom=<%=Server.URLEncode(SearchPriceFrom)%>&SearchPriceTo=<%=Server.URLEncode(SearchPriceTo)%>&SearchWaterfront=<%=Server.URLEncode(SearchWaterfront)%>&SearchSki=<%=Server.URLEncode(SearchSki)%>&SearchCondo=<%=Server.URLEncode(SearchCondo)%>&SearchResort=<%=Server.URLEncode(SearchResort)%>&SearchCountryClub=<%=Server.URLEncode(SearchCountryClub)%>&SearchFarmOrRanch=<%=Server.URLEncode(SearchFarmOrRanch)%>";
		//document.ContactBroker.submit();
		return true;
	}
}
function switchImage(num){
	var Pic1 = "<%=ListingProfile.Fields.Item("PicturePath1").Value%>";
	var Pic2 = "<%=ListingProfile.Fields.Item("PicturePath2").Value%>";
	var Pic3 = "<%=ListingProfile.Fields.Item("PicturePath3").Value%>";
	var Pic4 = "<%=ListingProfile.Fields.Item("PicturePath4").Value%>";
	var Pic5 = "<%=ListingProfile.Fields.Item("PicturePath5").Value%>";
	var Pic6 = "<%=ListingProfile.Fields.Item("PicturePath6").Value%>";
	var Pic7 = "<%=ListingProfile.Fields.Item("PicturePath7").Value%>";
	var Pic8 = "<%=ListingProfile.Fields.Item("PicturePath8").Value%>";
	document.ListingPic.src=eval("Pic"+num);
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
                <td bgcolor="#D6D4BA" width="10">&nbsp;</td>
                <td bgcolor="#D6D4BA" colspan="3" class="CastlesTextBody"><b>Property Detail</b></td>
              </tr>
              <tr> 
                <td bgcolor="#EDEBDB" width="10">&nbsp;</td>
                <td bgcolor="#EDEBDB" width="310" class="CastlesTextBody"> 
                  <p><br>
                    This is the details page for the listing you have requested 
                    to view.</p>
                </td>
                <td bgcolor="#EDEBDB" width="125" align="center"> 
                  <table width="100" cellspacing="0" cellpadding="2">
                    <tr> 
                      <td bgcolor="#D6D4BA" width="10" class="CastlesTextBody">&gt;</td>
                      <td bgcolor="#D6D4BA" class="CastlesTextBody"><b>&nbsp;<a href="default.asp" class="normal">New 
                        Search</a></b></td>
                    </tr>
                  </table>
                </td>
                <td bgcolor="#EDEBDB" width="150" align="center"> 
                  <table width="125" cellspacing="0" cellpadding="2">
                    <tr> 
                      <td bgcolor="#D6D4BA" width="10" class="CastlesTextBody">&gt;</td>
                      <td bgcolor="#D6D4BA" class="CastlesTextBody">
<%
if Featured = "Y" then
%>
					<b>&nbsp;<a href="/FeaturedListings/default.asp" class="normal">Featured Listings</a></b>
<%
else
%>					  
					  <b>&nbsp;<a href="searchResults.asp?SearchStateProvinceID=<%=Server.URLEncode(SearchStateProvinceID)%>&SearchCity=<%=Server.URLEncode(SearchCity)%>&SearchZipcode=<%=Server.URLEncode(SearchZipcode)%>&ZipListing=<%=Server.URLEncode(ZipListing)%>&SearchPriceFrom=<%=Server.URLEncode(SearchPriceFrom)%>&SearchPriceTo=<%=Server.URLEncode(SearchPriceTo)%>&SearchWaterfront=<%=Server.URLEncode(SearchWaterfront)%>&SearchSki=<%=Server.URLEncode(SearchSki)%>&SearchCondo=<%=Server.URLEncode(SearchCondo)%>&SearchResort=<%=Server.URLEncode(SearchResort)%>&SearchCountryClub=<%=Server.URLEncode(SearchCountryClub)%>&SearchFarmOrRanch=<%=Server.URLEncode(SearchFarmOrRanch)%>" class="normal">Back 
                        to Results</a></b>
<%
end if
%>						
						</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td colspan="4" height="1"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
              <tr> 
                <td colspan="4" height="1" bgcolor="#CCCCCC"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
              <tr> 
                <td colspan="4" height="1"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
            </table>
            <table width="595" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="595" bgcolor="#CC6600" valign="top" colspan="5" class="CastlesTextWhiteHeader" height="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
if ListingProfile.Fields.Item("ShowAddress").Value = "Y" then
%>				
				<%=ListingProfile.Fields.Item("Address").Value%>&nbsp;<%=ListingProfile.Fields.Item("Unit").Value%>
<%
end if
%>				
				</td>
              </tr>
              <tr> 
                <td valign="top" width="300"> 
<%
if ListingProfile.Fields.Item("PicturePath1").Value = "" or isNull(ListingProfile.Fields.Item("PicturePath1").Value) then
	PicturePath = "http://www.castlesmag.com/manage/images/noPic.gif"
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
	strMoreImages = "More images: <a href='javascript:switchImage(1)' class='normal'>1</a> "
	intcount = 2
	For i = 2 to 8
		If ListingProfile.Fields.Item("PicturePath" & i).Value = "" or isNull(ListingProfile.Fields.Item("PicturePath" & i).Value) then
			'nothing
		else
			strMoreImages = strMoreImages & " | <a href='javascript:switchImage("& i &")' class='normal'>" & intcount & "</a>"
			intcount = intcount + 1
		end if
	next
	
	Response.Write(strMoreImages)
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
                      <td colspan="2" height="1"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
                    </tr>
                  </table>
<%
if ContactSuccess <> "Y" then
%>				  
                  <table width="300" border="0" cellspacing="0" cellpadding="0">
                    <form name="ContactBroker" method="post" onSubmit="return Validate()">
                      <tr> 
					  	<td width="10" bgcolor="#D6D4BA">&nbsp;</td>
                        <td width="290" colspan="3" bgcolor="#D6D4BA" class="CastlesTextBody"><b>Contact 
                          a Broker about this property</b></td>
                      </tr>
                      <tr> 
                        <td width="300" colspan="4" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="80" bgcolor="#EDEBDB" class="CastlesTextBody">Name:</td>
                        <td width="190" bgcolor="#EDEBDB"> 
                          <input type="text" name="UserName" size="25" maxlength="100" class="CastlesTextBlack">
                        </td>
                        <td width="20" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="80" bgcolor="#EDEBDB" class="CastlesTextBody">Phone:</td>
                        <td width="190" bgcolor="#EDEBDB"> 
                          <input type="text" name="TelNumber" size="25" maxlength="100" class="CastlesTextBlack">
                        </td>
                        <td width="20" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="80" bgcolor="#EDEBDB" class="CastlesTextBody">Email:</td>
                        <td width="190" bgcolor="#EDEBDB"> 
                          <input type="text" name="EmailAddress" size="25" maxlength="100" class="CastlesTextBlack">
                        </td>
                        <td width="20" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="80" bgcolor="#EDEBDB" class="CastlesTextBody">Comments:</td>
                        <td width="190" bgcolor="#EDEBDB"> 
                          <textarea cols="25" rows="4" name="Comments" class="CastlesTextBlack"></textarea>
                        </td>
                        <td width="20" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                        <td width="80" bgcolor="#EDEBDB"><input type="hidden" name="ListingID" value="<%=ListingID%>"><input type="hidden" name="BrokerID" value="<%=ListingProfile.Fields.Item("BrokerID").Value%>"><input type="hidden" name="ListingAddress" value="<%=ListingProfile.Fields.Item("Address").Value%>"><input type="hidden" name="ListPrice" value="<%=ListingProfile.Fields.Item("ListPrice").Value%>"></td>
                        <td width="190" bgcolor="#EDEBDB" align="left" class="CastlesTextBody"><br><input type="submit" name="submit" value="Submit" class="CastlesTextBlack"></td>
                        <td width="20" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="300" colspan="4" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="300" colspan="4" bgcolor="#EDEBDB">&nbsp;</td>
                      </tr>
                    </form>
                  </table>
<%
else
%>
					<table width="300" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td width="10" bgcolor="#EDEBDB">&nbsp;</td>
							<td width="290" bgcolor="#EDEBDB" class="CastlesTextBody"><b>
								The Broker for this listing has been contacted successfully.<br><br>
								Thank You.</b>
							</td>
						</tr>
						<tr> 
                        	<td width="300" colspan="2" bgcolor="#EDEBDB">&nbsp;</td>
                      	</tr>
						  <tr> 
							<td width="300" colspan="2" bgcolor="#EDEBDB">&nbsp;</td>
						  </tr>
					</table>
<%
end if
%>
                </td>
                <td valign="top" width="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                <td valign="top" width="1" bgcolor="#CCCCCC"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
                <td valign="top" width="10">&nbsp;</td>
                <td valign="top" width="283"> 
					<table width="283" border="0" cellspacing="0" cellpadding="2">
						<tr> 
						  <td class="CastlesTextBody" align="left" width="150">
<%
if ListingProfile.Fields.Item("ShowAddress").Value = "Y" then
%>						  
						  <b><%=ListingProfile.Fields.Item("Address").Value%>&nbsp;<%=ListingProfile.Fields.Item("Unit").Value%> </b>
<%
end if
%>						  
						  </td>
							<td class="CastlesTextBody" align="right">
<%
if Not IsNull(ListingProfile.Fields.Item("ShowListPrice").Value) then
	if	 ListingProfile.Fields.Item("ShowListPrice").Value = "Y" then
%>								
							<b><%=FormatCurrency(ListingProfile.Fields.Item("ListPrice").Value,0)%></b>
							
<% else	%>
							<B>Price Upon Request</B>							
<%
	end if
end if	
%>							
							</td>
						</tr>
						<tr> 
						  <td class="CastlesTextBody" align="left" width="150">&nbsp;</td>
							<td class="CastlesTextBody" align="right">&nbsp;</td>
						</tr>
					</table>
                  <table width="283" border="0" cellspacing="0" cellpadding="2">
                   <tr> 
                      <td width="100" align="left" class="CastlesTextBody">ListingID:</td>
                      <td width="183" class="CastlesTextBlack"><%=ListingID%></td>
		  </tr>
                   <tr> 
                      <td width="100" align="left" class="CastlesTextBody">Address: 
                      </td>
                      <td width="183" class="CastlesTextBlack">
<%
if ListingProfile.Fields.Item("City").Value <> "0" then 
	Response.Write(ListingProfile.Fields.Item("City").Value & ",&nbsp;")
end if

if ListingProfile.Fields.Item("StateProvinceName").Value <> "International" then
	Response.Write(ListingProfile.Fields.Item("StateProvinceName").Value)
end if 

if ListingProfile.Fields.Item("Zipcode").Value <> 0 then 
	Response.Write("&nbsp;" & ListingProfile.Fields.Item("Zipcode").Value)
end if

if ListingProfile.Fields.Item("StateProvinceName").Value <> "International" and ListingProfile.Fields.Item("Zipcode").Value <> 0 then 
	Response.Write("<br>")
end if

if ListingProfile.Fields.Item("CountryName").Value <> "" then
	Response.Write(ListingProfile.Fields.Item("CountryName").Value)
end if
	
%></td>
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
                    <tr> 
                      <td colspan="2" align="left" class="CastlesTextBody" bgcolor="#EDEBDB">This Property Listing is provided by:</td>
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
                      <td width="100" align="left" class="CastlesTextBody" valign="top">Company: </td>
                      <td width="183" class="CastlesTextBlack"><%=CompanyName%></td>
                    </tr>
					<tr> 
                      <td width="100" align="left" class="CastlesTextBody" valign="top">Agent: </td>
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
					  Tel:  <%=TelNumber%><br>
					  <% if (Len(FaxNumber)<>0) then
					  		response.write "Fax:  " & FaxNumber & "<BR>"
						end if
					  %>
					  </td>
                    </tr>
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
