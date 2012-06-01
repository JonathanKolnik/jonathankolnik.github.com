<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/templates/castlesclientcnekt.asp" -->
<%

ImageNumber = Request.Cookies("ImageNumber")
If Len(ImageNumber) = 0 Then
	ImageNumber = 1
Else
	If ImageNumber = 3 Then
		ImageNumber = 1
	Else
		ImageNumber = ImageNumber + 1
	End If
End If
Response.Cookies("ImageNumber") = ImageNumber

'random featured property generator
Function getRandomNumber(upperbound)
	lowerbound=0
	Randomize
	getRandomNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
end Function

set Command1 = Server.CreateObject("ADODB.Command")
with command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_FeaturedProperty_getIDs"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	set FeaturedProperties = .Execute()
end with
set Command1 = nothing

FeaturedPropertiesArray = FeaturedProperties.getrows
FeaturedProperties.close
set FeaturedProperties = nothing
FeaturedPropertiesArrayNumRows=uBound(FeaturedPropertiesArray,2)

if(FeaturedPropertiesArrayNumRows>=2)then
	'response.write FeaturedPropertiesArrayNumRows
	index1=getRandomNumber(FeaturedPropertiesArrayNumRows)
	index2=getRandomNumber(FeaturedPropertiesArrayNumRows)
	while(index1=index2)
		index2 = getRandomNumber(FeaturedPropertiesArrayNumRows)
	wend
	index3 = getRandomNumber(FeaturedPropertiesArrayNumRows)
	While((index1=index3)or(index2=index3))
		index3 = getRandomNumber(FeaturedPropertiesArrayNumRows)
	Wend
	id1 = FeaturedPropertiesArray(0,index1)
	id2 = FeaturedPropertiesArray(0,index2)
	id3 = FeaturedPropertiesArray(0,index3)
	
	set Command1 = Server.CreateObject("ADODB.Command")
	with command1
		.ActiveConnection = connect
		.CommandText = "Castles_ClientSide_FeaturedProperty_Display"
		.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
		.Parameters.Append .CreateParameter("@FeaturedListing1",200,1,200,id1)
		.Parameters.Append .CreateParameter("@FeaturedListing2",200,1,200,id2)
		.Parameters.Append .CreateParameter("@FeaturedListing3",200,1,200,id3)	
		.CommandType = 4
		.CommandTimeout = 0
		.Prepared = true
		set FeaturedProperties = .Execute()
	end with
	set Command1 = nothing
	
	FeaturedPropertiesArray = FeaturedProperties.getrows
	FeaturedProperties.close
	set FeaturedProperties = nothing
	FeaturedPropertiesArrayNumRows=uBound(FeaturedPropertiesArray,2)
'	for columncount=0 to FeaturedPropertiesArrayNumRows
'		response.write FeaturedPropertiesArray(columncount,0)&"<BR>"
'	next
	for rowcount=0 to FeaturedPropertiesArrayNumRows
		'response.write "FeaturedListingID is: "&FeaturedPropertiesArray(0,rowcount)&"<BR>"
		if(FeaturedPropertiesArray(0,rowcount)=id1)then
			title1 = FeaturedPropertiesArray(1,rowcount)
			desc1 = FeaturedPropertiesArray(2,rowcount)
			pic1 = FeaturedPropertiesArray(3,rowcount)
		end if
		if(FeaturedPropertiesArray(0,rowcount)=id2)then
			title2 = FeaturedPropertiesArray(1,rowcount)
			desc2 = FeaturedPropertiesArray(2,rowcount)
			pic2 = FeaturedPropertiesArray(3,rowcount)
		end if
		if(FeaturedPropertiesArray(0,rowcount)=id3)then
			title3 = FeaturedPropertiesArray(1,rowcount)
			desc3 = FeaturedPropertiesArray(2,rowcount)
			pic3 = FeaturedPropertiesArray(3,rowcount)
		end if	
	next
end if
%>
<html>
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
	daughter = window.open("http://216.119.112.46/manage/default.asp?Broker=Y",'daughter','toolbar=no,scrollbars=yes,resizable');
}
function doSearch(){
	if((document.PropertyQuickSearch.SearchStateProvinceID.value == 0)&&(document.PropertyQuickSearch.SearchZipcode.value == "")){
		alert("Please select a Location or enter a Zipcode.");
	}else{
		document.PropertyQuickSearch.submit();
	}
}
//-->
</script>

</head>
<body bgcolor="#FFFFFF" leftmargin="3" topmargin="3" marginwidth="3" marginheight="3">
<table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#333399" height="10" width="350"><img src="images/clear10pixel.gif" width="10" height="10"></td>
		  <td width="400" height="10"><img src="images/tagline.gif" width="387" height="10"></td>
        </tr>
        <tr> 
          <td width="350">
            <table width="300" border="0" cellspacing="0" cellpadding="8">
              <tr>
                <td><img src="images/castlemed.gif" width="151" height="91"></td>
              </tr>
            </table>
          </td>
		  <td width="400" valign="bottom" align="right">
		  	<table width="400" cellspacing="0" cellpadding="0" border="0">
				<tr>
					
                <td width="300" align="right" class="CastlesTextBodyBold"><a href="javascript:BrokerLogin()" class="black">&gt; 
                  Submit FREE Online Listing/Magazine Ad (Brokers)</a></td>
                <td width="100" align="right" class="CastlesTextBodyBold"><a href="contactcastles/" class="black">&gt; 
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
                <td colspan="3" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#A0A0A0"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="Listings/default.asp" class="black">Search 
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
                <td colspan="3" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="/aboutcastles/" class="black">About 
                        Castles</a></td>
                    </tr>
                    <tr> 
                      <td class="CastlesTextNav"><a href="contactcastles/" class="black">Contact 
                        Castles</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="/subscribetomagazine/" class="black">Subscribe 
                        to Magazine</a></td>
                    </tr>
                  </table>
                </td>
                <td width="10">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="3" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#BABABA"> 
                <td width="10">&nbsp;</td>
                <td width="130"> 
                  <table width="130" border="0" cellspacing="0" cellpadding="3">
                    <tr> 
                      <td class="CastlesTextNav"><a href="/openbrokeraccount/" class="black">Open 
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
                      <td class="CastlesTextNav"><a href="/advertisingrates/" class="black">Advertising 
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
                      <td class="CastlesTextNav"><a href="/submissionguidelines/" class="black">Submission 
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
            </table>
          </td>
          <td width="2"><img src="images/clear10pixel.gif" width="2" height="1"></td>
          <td bgcolor="#CCCCCC" width="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td width="2"><img src="images/clear10pixel.gif" width="2" height="1"></td>
          <td bgcolor="#EDEBDB" width="196" valign="top"> 
            <table width="196" border="0" cellspacing="0" cellpadding="0">
              <form name="PropertyQuickSearch" method="post" action="/Listings/searchResults.asp">
                <tr> 
                  <td width="10" bgcolor="#D6D4BA">&nbsp;</td>
                  <td width="176" bgcolor="#D6D4BA"> 
                    <table width="130" border="0" cellspacing="0" cellpadding="3" height="20">
                      <tr> 
                        <td class="CastlesTextBodyBold">Property Quick-find</td>
                      </tr>
                    </table>
                  </td>
                  <td width="10" bgcolor="#D6D4BA">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                </tr>
                <tr> 
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                  <td width="176" class="CastlesTextBody">By Location:</td>
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                </tr>
                <tr> 
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                  <td width="176" class="CastlesTextBlack"> 
                    <select name="SearchStateProvinceID" class="CastlesTextBlack">
						<option value="0">Select a Location</option>
<%
set Command1 = Server.CreateObject("ADODB.Command")
with command1
	.ActiveConnection = connect
	.CommandText = "Castles_ClientSide_StateProvinceAndCountry_ABV_DynamicList"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = true
	set StateCountry = .Execute()
end with
set Command1 = nothing
While Not StateCountry.EOF
%>					
                    	<option value="<%=StateCountry.Fields.Item("StateProvinceID").Value%>"><%=StateCountry.Fields.Item("StateProvinceAbv").Value%> / <%=StateCountry.Fields.Item("CountryAbv").Value%></option>  
<%
	StateCountry.MoveNext()
Wend
StateCountry.close
%>						
                    </select>
                  </td>
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                </tr>
                <tr> 
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                  <td width="176" class="CastlesTextBodyBold">Or </td>
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                </tr>
                <tr> 
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                  <td width="176" class="CastlesTextBody">By Zipcode:</td>
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" height="2" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="2"></td>
                </tr>
                <tr> 
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                  <td width="176" class="CastlesTextBlack"> 
                    <input type="text" name="SearchZipcode" class="CastlesTextBlack">
                  </td>
                  <td width="10" bgcolor="#EDEBDB"><input type="hidden" name="SearchPriceFrom" value="0"><input type="hidden" name="SearchPriceTo" value="1000000000"></td>
                </tr>
                <tr bgcolor="#EDEBDB"> 
                  <td colspan="3" height="10"><img src="images/clear10pixel.gif" width="1" height="10"></td>
                </tr>
                <tr> 
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                  <td width="176"> 
                    <table width="35" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td class="CastlesTextBodyBold" align="center" bgcolor="#D6D4BA"><input type="button" name="Submit" value="Go" class="CastlesTextBlack" onClick="doSearch()"></td>
                      </tr>
                    </table>
                  </td>
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" height="10" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="10"></td>
                </tr>
                <tr> 
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                  <td width="176" class="CastlesTextBody">You can also search 
                    more specifically for properties on the Castles website by 
                    using our <a href="Listings/default.asp" class="normal">Advanced 
                    Search</a> option.</td>
                  <td width="10" bgcolor="#EDEBDB">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="3" height="10" bgcolor="#EDEBDB"><img src="images/clear10pixel.gif" width="1" height="10"></td>
                </tr>
              </form>
            </table>
          </td>
          <td width="2"><img src="images/clear10pixel.gif" width="2" height="1"></td>
          <td bgcolor="#CCCCCC" width="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td width="396" valign="top"> 
            <table width="387" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
                <td rowspan="5" bgcolor="#FFFFFF"><a href="subscribetomagazine/default.asp"><img src="images/image_<%=ImageNumber%>.jpg" width="387" height="158" border="0" href="#" alt="Click Here Subscribe to Castles Today and get 50% off."></a></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
                <td bgcolor="#FFFFFF" height="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
              </tr>
              <tr> 
                <td bgcolor="#CCCCCC" width="2"><img src="images/clear10pixel.gif" width="2" height="1"></td>
                <td bgcolor="#CCCCCC" height="1"><img src="images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
                <td bgcolor="#FFFFFF" height="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
              </tr>
              <tr> 
                <td bgcolor="#FFFFFF" width="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
                <td> 
                  <table width="387" border="0" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#EDEBDB"> 
                      <td width="10" height="10"><img src="images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="329" height="10"><img src="images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                    <tr bgcolor="#EDEBDB"> 
                      <td width="10">&nbsp;</td>
                      <td width="329" class="CastlesTextBody"><span class="CastlesTextBodyBold"><a href="subscribetomagazine/default.asp" class="normal">Click 
                        Here Subscribe to Castles Today and get 50% off.</a></span><br>
                        There has never been a better time to subscribe to Castles 
                        Magazine and for a limited time only we are willing to 
                        give you a discount. <span class="CastlesTextBodyBold">[&gt;<a href="subscribetomagazine/default.asp" class="normal">More</a>]<br>
                        <br>
                        </span></td>
                    </tr>
                    <tr bgcolor="#EDEBDB"> 
                      <td width="10" height="10"><img src="images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="329" height="10"><img src="images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td width="150" valign="top" bgcolor="#FFFFFF" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#FFFFFF" width="196" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2"><img src="images/clear10pixel.gif" width="2" height="2"></td>
          <td bgcolor="#CCCCCC" width="1" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td bgcolor="#FFFFFF" width="397" height="2"><img src="images/clear10pixel.gif" width="1" height="2"></td>
        </tr>
        <tr> 
          <td bgcolor="#CCCCCC" width="150" valign="top" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="196" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="2" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="1" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
          <td bgcolor="#CCCCCC" width="396" height="1"><img src="images/clear10pixel.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td width="150" valign="top" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="2" height="2"></td>
          <td width="196" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td width="2" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="2" height="2"></td>
          <td width="1" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="1" height="2"></td>
          <td width="396" height="2" bgcolor="#FFFFFF"><img src="images/clear10pixel.gif" width="1" height="2"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign="top"> 
      <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="154">&nbsp;</td>
          <td width="595" bgcolor="#333399"> 
            <table width="595" border="0" cellspacing="0" cellpadding="3" height="25">
              <tr> 
                <td width="290" class="CastlesTextWhiteHeader">&nbsp;&nbsp;Featured 
                  Properties</td>
                <td width="293" class="CastlesTextWhiteBold" align="right">&gt; 
                  <a href="/FeaturedListings/default.asp" class="white">See More Featured Properties</a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td width="154">&nbsp;</td>
          <td width="595" valign="top"> 
            <table width="595" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="100"><a href="/Listings/ListingDetails.asp?Featured=Y&ListingID=<%=id1%>"><img src="<%=pic1%>" width="100" height="100" border="0" onError=this.src="/manage/images/noPic.gif"></a></td>
                <td width="495"> 
                  <table width="495" border="0" cellspacing="0" cellpadding="0" height="75">
                    <tr bgcolor="#EDEBDB"> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="280" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="195" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                    <tr bgcolor="#EDEBDB"> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" bgcolor="#FFFFFF">&nbsp;</td>
                      <td width="280" class="CastlesTextBody" bgcolor="#FFFFFF"><span class="CastlesTextBodyBold"><%=title1%></span><br>
                        <%=desc1%> <span class="CastlesTextBodyBold">[&gt;<a href="/Listings/ListingDetails.asp?Featured=Y&ListingID=<%=id1%>" class="normal">More</a>]</span></td>
                      <td width="195" bgcolor="#FFFFFF">&nbsp;</td>
                    </tr>
                    <tr bgcolor="#EDEBDB"> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="280" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="195" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td colspan="2" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#CCCCCC"> 
                <td colspan="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td colspan="2" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr> 
                <td width="100"><a href="/Listings/ListingDetails.asp?Featured=Y&ListingID=<%=id2%>"><img src="<%=pic2%>" width="100" height="100" border="0" onError=this.src="/manage/images/noPic.gif"></a></td>
                <td width="495"> 
                  <table width="495" border="0" cellspacing="0" cellpadding="0" height="75">
                    <tr> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="280" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="195" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                    <tr> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" bgcolor="#FFFFFF">&nbsp;</td>
                      <td width="280" class="CastlesTextBody"><span class="CastlesTextBodyBold"><%=title2%></span><br>
                        <%=desc2%><span class="CastlesTextBodyBold">[&gt;<a href="/Listings/ListingDetails.asp?Featured=Y&ListingID=<%=id2%>" class="normal">More</a>]</span></td>
                      <td width="195" bgcolor="#FFFFFF">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="280" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="195" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td colspan="2" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#CCCCCC"> 
                <td colspan="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td colspan="2" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr> 
                <td width="100"><a href="/Listings/ListingDetails.asp?Featured=Y&ListingID=<%=id3%>"><img src="<%=pic3%>" width="100" height="100" border="0" onError=this.src="/manage/images/noPic.gif"></a></td>
                <td width="495"> 
                  <table width="495" border="0" cellspacing="0" cellpadding="0" height="75">
                    <tr bgcolor="#EDEBDB"> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="280" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="195" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                    <tr bgcolor="#EDEBDB"> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" bgcolor="#FFFFFF">&nbsp;</td>
                      <td width="280" class="CastlesTextBody" bgcolor="#FFFFFF"><span class="CastlesTextBodyBold"><%=title3%></span><br>
                        <%=desc3%><span class="CastlesTextBodyBold">[&gt;<a href="/Listings/ListingDetails.asp?Featured=Y&ListingID=<%=id3%>" class="normal">More</a>]</span></td>
                      <td width="195" bgcolor="#FFFFFF">&nbsp;</td>
                    </tr>
                    <tr bgcolor="#EDEBDB"> 
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="1" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="2" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="2" height="10"></td>
                      <td width="10" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="280" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                      <td width="195" height="10" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="1" height="10"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td colspan="2" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
              <tr bgcolor="#CCCCCC"> 
                <td colspan="2" height="1"><img src="../images/clear10pixel.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td colspan="2" height="2"><img src="../images/clear10pixel.gif" width="1" height="2"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td width="154">&nbsp;</td>
          <td width="595" class="CastlesTextBody">&copy; 2003 <a href="http://www.castlesmag.com" class="normal">Castles 
            Magazine</a> &nbsp;&nbsp;All rights reserved. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="CastlesTextBodyBold">&gt;<a href="/misc/privacypolicy.asp" class="normal">Privacy 
            Policy</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&gt;<a href="/misc/termsofuse.asp" class="normal">Terms 
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
</html>
