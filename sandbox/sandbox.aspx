<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="UTF-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="templates/castlesclientcnekt.asp" -->
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
			price1 = FeaturedPropertiesArray(4,rowcount)
			ShowListPrice1 = FeaturedPropertiesArray(5,rowcount)
		end if
		if(FeaturedPropertiesArray(0,rowcount)=id2)then
			title2 = FeaturedPropertiesArray(1,rowcount)
			desc2 = FeaturedPropertiesArray(2,rowcount)
			pic2 = FeaturedPropertiesArray(3,rowcount)
			price2 = FeaturedPropertiesArray(4,rowcount)
			ShowListPrice2 = FeaturedPropertiesArray(5,rowcount)
		end if
		if(FeaturedPropertiesArray(0,rowcount)=id3)then
			title3 = FeaturedPropertiesArray(1,rowcount)
			desc3 = FeaturedPropertiesArray(2,rowcount)
			pic3 = FeaturedPropertiesArray(3,rowcount)
			price3 = FeaturedPropertiesArray(4,rowcount)
			ShowListPrice3 = FeaturedPropertiesArray(5,rowcount)
		end if	
	next
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" type="text/css" href="css/css-slider.css" />

<script type="text/javascript">window.location.hash = '#image-1'</script>
<title>Untitled Document</title>
<script language="JavaScript">
<!--
function BrokerLogin(){
	daughter = window.open("http://www.castlesmag.com/manage/default.asp?Broker=Y",'daughter','toolbar=Yes,width=600,height=400,top=50,left=50,scrollbars=yes,resizable');
}
function Isnumeric(field) {
var valid = "0123456789"
var ok = "yes";
var temp;
for (var i=0; i<field.value.length; i++) {
temp = "" + field.value.substring(i, i+1);
if (valid.indexOf(temp) == "-1") ok = "no";
}
if (ok == "no") {
//alert("Zip Code may only contain numbers.");
	if (document.PropertyQuickSearch.ZipListing[0].checked) 
			alert("Zip Code may only contain numbers.");
	else 
			alert("Listing ID may only contain numbers.");
field.focus();
field.select();
return false;
   }
 return true; 
}

function doSearch(){
	
	//if((document.PropertyQuickSearch.SearchStateProvinceID.value == 0)&&(document.PropertyQuickSearch.SearchZipcode.value == "")){
		//alert("Please select a Location or enter a Zipcode.");
		//return false;
//	}else{
	if (document.PropertyQuickSearch.SearchStateProvinceID.value == "" && document.PropertyQuickSearch.SearchZipcode.value == "") 
		{
		alert("Please select a Location or enter a Zip Code or enter a Listing ID.");
		return false;
		}
	else if (document.PropertyQuickSearch.SearchStateProvinceID.value != "")
		{
			document.PropertyQuickSearch.submit();
		}
	else if ((document.PropertyQuickSearch.SearchZipcode.value != "") && Isnumeric(document.PropertyQuickSearch.SearchZipcode)) 
		{
			/*Isnumeric(document.PropertyQuickSearch.SearchZipcode);
			return false;*/
			document.PropertyQuickSearch.submit();
		}	
	
}

</script>

</head>
<center>
<body>
	<table width="1000" colspan="3">
    	<tr>
        	
        	<td align="right" colspan="3">
            <a href="javascript:BrokerLogin()" class="black">Broker 
                        Login</a>             
            </td>
         </tr>
         <tr>
         	<td>
            	<img src="images/castleslogo.gif" alt="logo" height="76" width="209" >
            </td>
         </tr>
         <tr>
         	<td width="425" valign="top" >
            	<table>
                     <tr>
                        <td bgcolor="#46443a" width="425" height="25" style="font:'Open Sans Semibold'; color:#FFF; font-size:14px; padding-left:5px">
                          Property Quick-find
                        </td>
                     </tr>
                     <tr>
                     	<td>
                        
                        
                        </td>
                     </tr>
                 </table>
                        
             </td>           
            <td width="5px">
            
            </td>
            <td>
            <div id="slider">

	<div id="image-1">
		<a href=""><img src="s1.jpeg" alt="" /></a>
		<a class="slider-nav" href="#image-1"></a>
	</div>
	<div id="image-2">
		<a href=""><img src="s2.jpeg" alt="" /></a>
		<a class="slider-nav" href="#image-2"></a>
	</div>
	<div id="image-3">
		<a href=""><img src="s3.jpeg" alt="" /></a>
		<a class="slider-nav" href="#image-3"></a>
	</div>
	<div id="image-4">
		<a href=""><img src="s4.jpeg" alt="" /></a>
		<a class="slider-nav" href="#image-4"></a>
	</div>
	
	
</div>

            </td>
         </tr>  
          
            







	</table>




</body>
<center>
</html>
