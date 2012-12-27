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

CountryID = findValue("CountryID")
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

ZipListing = findValue("ZipListing")


if countryID = "" then
	countryID = 0
end if


' for pagination stuff at bottom
RecsPerPage = 10

Page = Request.QueryString("Page")
If Len(Page) = 0 Then
	Page = 1
End If

TotalRecords = Request.QueryString("TotalRecords")
MoreRecords = Request.QueryString("MoreRecords")
' end pagination stuff

'Generates Appropriate Pagination for Simple System Search Footer
Function DCSystemSimpleSearchPagination(Page,WordPhrase_JumpToPage,PageName,MoreRecords,TotalRecords,RecsPerPage,CountryID,StateProvinceID,City,Zipcode,PriceFrom,PriceTo,Waterfront,Ski,Condo,Resort,CountryClub,FarmOrRanch)
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
	DCSystemSimpleSearchPagination = "Jump To Page: "

	For i = 1 to Paginations 
		If cStr(PageCount) <> cStr(Page) Then
			DCSystemSimpleSearchPagination = DCSystemSimpleSearchPagination &  "<a href=" & DQ & PageName & "?CountryID=" & Server.URLEncode(CountryID) & "&SearchStateProvinceID=" & Server.URLEncode(SearchStateProvinceID) & "&SearchCity=" & Server.URLEncode(SearchCity) & "&SearchZipcode=" & Server.URLEncode(SearchZipcode) & "&ZipListing=" & Server.URLEncode(ZipListing) & "&SearchPriceFrom=" & Server.URLEncode(SearchPriceFrom) & "&SearchPriceTo=" & Server.URLEncode(SearchPriceTo) & "&SearchWaterfront" & Server.URLEncode(SearchWaterfront) & "&SearchSki=" & Server.URLEncode(SearchSki) & "&SearchCondo=" & Server.URLEncode(SearchCondo) & "&SearchResort=" & Server.URLEncode(SearchResort) & "&SearchCountryClub=" & Server.URLEncode(SearchCountryClub) & "&SearchFarmOrRanch=" & Server.URLEncode(SearchFarmOrRanch) & "&Page=" & PageCount & "&TotalRecords=" & TotalRecords & DQ &" class=""orange""><b>" & PageCount & "</b></a>|"
		Else
			DCSystemSimpleSearchPagination = DCSystemSimpleSearchPagination & "<b>" & PageCount & "</b>|" 
		End If
		PageCount = PageCount + 1
	Next
End Function


If SearchZipcode <> "" then
	if isnumeric(SearchZipcode) = true then
		if ZipListing = "zipcode" then	
			Whereclause = "WHERE Castles_Listings.Zipcode=" & SearchZipcode & " AND Castles_Listings.ListPrice>=" & SearchPriceFrom & " AND Castles_Listings.ListPrice<=" & SearchPriceTo
		else
			Whereclause = "WHERE Castles_Listings.ListingID=" & SearchZipcode & " AND Castles_Listings.ListPrice>=" & SearchPriceFrom & " AND Castles_Listings.ListPrice<=" & SearchPriceTo
		end if
	else
		Whereclause = "WHERE Castles_Listings.ListPrice>=" & SearchPriceFrom & " AND Castles_Listings.ListPrice<=" & SearchPriceTo
	end if
	
	if SearchWaterfront <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Waterfront='Y'"
	end if
	if SearchSki <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Ski='Y'"
	end if
	if SearchCondo <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Condo='Y'"
	end if
	if SearchResort <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Resort='Y'"
	end if
	if SearchCountryClub <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.CountryClub='Y'"
	end if
	if SearchFarmOrRanch <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.FarmOrRanch='Y'"
	end if
End if
If SearchZipcode = "" then
	Whereclause = "WHERE Castles_Listings.Active = 'Y'"
	if CountryID <> 0   then
		Whereclause = Whereclause & " AND Castles_Listings.CountryID=" & CountryID & "" 
	end if
	if SearchPriceFrom <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.ListPrice>=" & SearchPriceFrom & " AND Castles_Listings.ListPrice<=" & SearchPriceTo & ""
	end if
	'Response.write (SearchStateProvinceID)
	if SearchStateProvinceID <> 0 then
		if cint(SearchStateProvinceID) = -1 then
			Whereclause = Whereclause & " AND Castles_Listings.StateProvinceID = 0"
		else
			Whereclause = Whereclause & " AND Castles_Listings.StateProvinceID=" & SearchStateProvinceID & ""
		end if
	end if
	if SearchCity <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.City LIKE '" & SearchCity & "%'"
	end if
	if SearchWaterfront <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Waterfront='Y'"
	end if
	if SearchSki <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Ski='Y'"
	end if
	if SearchCondo <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Condo='Y'"
	end if
	if SearchResort <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.Resort='Y'"
	end if
	if SearchCountryClub <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.CountryClub='Y'"
	end if
	if SearchFarmOrRanch <> "" then
		Whereclause = Whereclause & " AND Castles_Listings.FarmOrRanch='Y'"
	end if
End if
'response.write "Whereclause: " & Whereclause & "<BR>"

Set Command1 = Server.CreateObject("ADODB.Command")
With Command1	
	.ActiveConnection = Connect
	.CommandText = "Castles_ClientSide_ListingsSearch"
	.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
	.Parameters.Append .CreateParameter("@Page", 200, 1,200,Page)
	.Parameters.Append .CreateParameter("@RecsPerPage", 200, 1,200,RecsPerPage)
	.Parameters.Append .CreateParameter("@Whereclause",200,1,20000,Whereclause)
	.CommandType = 4
	.CommandTimeout = 0
	.Prepared = True
	Set ListingsSearch = .Execute()
End With
Set Command1 = Nothing
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
	}else{
		document.ContactBroker.submit();
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
                <td bgcolor="#D6D4BA" width="10">&nbsp;</td>
                <td bgcolor="#D6D4BA" colspan="3" class="CastlesTextBody"><b>Search 
                  Results </b></td>
              </tr>
              <tr> 
                <td bgcolor="#EDEBDB" width="10">&nbsp;</td>
                <td bgcolor="#EDEBDB" width="310" class="CastlesTextBody"> 
                  <p><br>
                    For more information on a listing, click on the 'View Details' 
                    link or the photo</p>
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
                <td bgcolor="#EDEBDB" width="150" align="center">&nbsp; </td>
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
                <td width="595" bgcolor="#CC6600" valign="top" colspan="5" class="CastlesTextWhiteHeader">&nbsp;</td>
              </tr>
<%
ListingsHasRecords = ""
If ListingsSearch.EOF then
	ListingsHasRecords = false
%>
			<tr>
				<td width="20">&nbsp;</td>
				<td width="575" colspan="4" class="CastlesTextBody">No Results Found. . . Please Search Again</td>
			</tr>	  
<%
End if
If NOT ListingsSearch.EOF then
	ListingsHasRecords = true
	SearchResultsReturned = ListingsSearch.getrows
	Count = 1
	SearchResultsArrayNumRows = uBound(SearchResultsReturned,2)

	Field_ID = 0
	Field_ListingID = 1
	Field_Address = 2
	Field_Unit = 3
	Field_City = 4
	Field_StateProvinceName = 5
	Field_ZipCode = 6
	Field_Bedrooms = 7
	Field_FullBaths = 8
	Field_HalfBaths = 9
	Field_ListPrice = 10
	Field_PicturePath1 = 11
	Field_ShowAddress = 12
	Field_ShowListPrice = 13
	Field_CompanyName = 14
	Field_FirstName = 15
	Field_MiddleInitial = 16
	Field_LastName = 17
	Field_TelNumber = 18
	Field_CountryName = 19
	Field_MoreRecords = 20

	For SearchResultsArrayRowCounter = 0 to SearchResultsArrayNumRows
		ListingID = SearchResultsReturned(Field_ListingID,SearchResultsArrayRowCounter)
		Address = SearchResultsReturned(Field_Address,SearchResultsArrayRowCounter)
		Unit = SearchResultsReturned(Field_Unit,SearchResultsArrayRowCounter)
		City = SearchResultsReturned(Field_City,SearchResultsArrayRowCounter)
		StateProvinceName = SearchResultsReturned(Field_StateProvinceName,SearchResultsArrayRowCounter)
		ZipCode = SearchResultsReturned(Field_ZipCode,SearchResultsArrayRowCounter)
		CountryName = SearchResultsReturned(Field_CountryName,SearchResultsArrayRowCounter)
		Bedrooms = SearchResultsReturned(Field_Bedrooms,SearchResultsArrayRowCounter)
		FullBaths = SearchResultsReturned(Field_FullBaths,SearchResultsArrayRowCounter)
		HalfBaths = SearchResultsReturned(Field_HalfBaths,SearchResultsArrayRowCounter)
		ListPrice = SearchResultsReturned(Field_ListPrice,SearchResultsArrayRowCounter)
		PicturePath1 = SearchResultsReturned(Field_PicturePath1,SearchResultsArrayRowCounter)
		ShowAddress = SearchResultsReturned(Field_ShowAddress,SearchResultsArrayRowCounter)
		ShowListPrice = SearchResultsReturned(Field_ShowListPrice,SearchResultsArrayRowCounter)
		CompanyName = SearchResultsReturned(Field_CompanyName,SearchResultsArrayRowCounter)
		FirstName = SearchResultsReturned(Field_FirstName,SearchResultsArrayRowCounter)
		MiddleInitial = SearchResultsReturned(Field_MiddleInitial,SearchResultsArrayRowCounter)
		LastName = SearchResultsReturned(Field_LastName,SearchResultsArrayRowCounter)
		TelNumber = SearchResultsReturned(Field_TelNumber,SearchResultsArrayRowCounter)
		MoreRecords = SearchResultsReturned(Field_MoreRecords,SearchResultsArrayRowCounter)
		
		if PicturePath1 = "" or isNull(PicturePath1) then
			PicturePath = "http://www.castlesmag.com/manage/images/noPic.gif"
		else
			PicturePath = PicturePath1
		end if	
%>
			  <tr> 
                <td width="100" valign="top"><a href="ListingDetails.asp?ListingID=<%=ListingID%>&SearchStateProvinceID=<%=Server.URLEncode(SearchStateProvinceID)%>&SearchCity=<%=Server.URLEncode(SearchCity)%>&SearchZipcode=<%=SearchZipcode%>&ZipListing=<%=ZipListing%>&SearchPriceFrom=<%=Server.URLEncode(SearchPriceFrom)%>&SearchPriceTo=<%=Server.URLEncode(SearchPriceTo)%>&SearchWaterfront=<%=Server.URLEncode(SearchWaterfront)%>&SearchSki=<%=Server.URLEncode(SearchSki)%>&SearchCondo=<%=Server.URLEncode(SearchCondo)%>&SearchResort=<%=Server.URLEncode(SearchResort)%>&SearchCountryClub=<%=Server.URLEncode(SearchCountryClub)%>&SearchFarmOrRanch=<%=Server.URLEncode(SearchFarmOrRanch)%>" class="normal"><img src="<%=PicturePath%>" width="100" height="100" border="0"></a></td>
				<td width="15">&nbsp;</td>
				<td width="325" valign="top" class="CastlesTextBody">
					<br>
<% if Not IsNull(ShowListPrice)	then
	if ShowListPrice = "Y"	then			
%>					
						<B><%=FormatCurrency(ListPrice,0)%></B>
<%	else %>
						<B>Priced Upon Request</B>						
<% end if
end if %>						
						<br><br>
<%
if ShowAddress = "Y" then
	if Unit <> "" then
		Unit = ", " & Unit
	end if
%>					
					<%=Address%><%=Unit%><br>
<%
end if
%>					
					<% if City <> "0" then 
						Response.Write(City & ",")
					   end if
					%>

					<% if StateProvinceName <> "International" then
						Response.Write(StateProvinceName)
						end if 
							%> 
						
					<% if ZipCode <> 0 then 
							Response.Write(ZipCode)
						end if
						if StateProvinceName <> "International" and ZipCode <> 0 then 
							Response.Write(",")
						end if

						if CountryName <> "" then
							Response.Write(CountryName)
						end if
					%><br><br>

					Bedrooms: <%=Bedrooms%><br>
					Full Baths/Half Baths: <%=FullBaths%>/<%=HalfBaths%> <br>
                  <br>
                  Listing Provided By:<br>
				  <%=CompanyName%> - <%=LastName%>,&nbsp;<%=FirstName%>&nbsp;<%=MiddleInitial%>&nbsp;&nbsp;tel:&nbsp;<%=TelNumber%>
                </td>
				<td width="15">&nbsp;</td>
				<td width="140" valign="top" class="CastlesTextBody">
					<br><a href="ListingDetails.asp?ListingID=<%=ListingID%>&SearchStateProvinceID=<%=Server.URLEncode(SearchStateProvinceID)%>&SearchCity=<%=Server.URLEncode(SearchCity)%>&SearchZipcode=<%=SearchZipcode%>&ZipListing=<%=ZipListing%>&SearchPriceFrom=<%=Server.URLEncode(SearchPriceFrom)%>&SearchPriceTo=<%=Server.URLEncode(SearchPriceTo)%>&SearchWaterfront=<%=Server.URLEncode(SearchWaterfront)%>&SearchSki=<%=Server.URLEncode(SearchSki)%>&SearchCondo=<%=Server.URLEncode(SearchCondo)%>&SearchResort=<%=Server.URLEncode(SearchResort)%>&SearchCountryClub=<%=Server.URLEncode(SearchCountryClub)%>&SearchFarmOrRanch=<%=Server.URLEncode(SearchFarmOrRanch)%>" class="normal">View Details</a>
				</td>
			</tr>
			<tr> 
                <td colspan="5" height="1"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
              <tr> 
                <td colspan="5" height="1" bgcolor="#CCCCCC"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
              <tr> 
                <td colspan="5" height="1"><img src="../images/clear10pixel.gif" width="2" height="1"></td>
              </tr>
<%
	Next
End if
%>
			</table>
			<table width="595" cellspacing="0" cellpadding="0" border="0">
			<tr>
				<td width="20">&nbsp;</td>
				<td width="575" class="CastlesTextBody">
<%
'Generates Appropriate Pagination Code
if ListingsHasRecords then
'If Len(SQLStmt) <> 0 Then
	PageName = "searchResults.asp"
	TotalRecords = (MoreRecords)+(RecsPerPage)
	PaginationCode = DCSystemSimpleSearchPagination(Page,WordPhrase_JumpToPage,PageName,MoreRecords,TotalRecords,RecsPerPage,CountryID,SearchStateProvinceID,SearchCity,SearchZipcode,SearchPriceFrom,SearchPriceTo,SearchWaterfront,SearchSki,SearchCondo,SearchResort,SearchCountryClub,SearchFarmOrRanch)
	Response.Write PaginationCode
End If
%>		
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
				
          <td width="595"><!-- #BeginEditable "subBody" --><%=SearchStateProvinceID%><!-- #EndEditable --></td>
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
