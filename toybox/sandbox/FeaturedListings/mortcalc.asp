<%
ListPrice = Request.QueryString("ListPrice")




%>
<html>
<head>
<title>Castles Unlimited - Mortgage Calculator</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
<!--
.CastlesTextBlack {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBlackBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 13px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}
.CastlesTextBody {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #666666; text-decoration: none}
.CastlesTextBodyBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 13px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #666666; text-decoration: none}
.CastlesTextGray {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #666666; text-decoration: none}
.CastlesTextGrayBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 13px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #666666; text-decoration: none}
.CastlesTextWhite {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextWhiteBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 13px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FFFFFF; text-decoration: none}
.CastlesTextRed {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FF0033; text-decoration: none}
.CastlesTextRedBig {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 9px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #FF0033; text-decoration: none}

A.red:link    { text-decoration: none; color: "#FF0033"; }
A.red:visited { text-decoration: none; color: "#FF0033"; }
A.red:active  { text-decoration: none; color: "#FF0033"; }
A.red:hover   { text-decoration: underline; color: "#FF0033"; }

A.white:link    { text-decoration: none; color: "#FFFFFF"; }
A.white:visited { text-decoration: none; color: "#FFFFFF"; }
A.white:active  { text-decoration: none; color: "#FFFFFF"; }
A.white:hover   { text-decoration: underline; color: "#FFFFFF"; }

A.gray:link    { text-decoration: none; color: "#666666"; }
A.gray:visited { text-decoration: none; color: "#666666"; }
A.gray:active  { text-decoration: none; color: "#666666"; }
A.gray:hover   { text-decoration: underline; color: "#666666"; }

A.lightgray:link    { text-decoration: none; color: "#CCCCCC"; }
A.lightgray:visited { text-decoration: none; color: "#CCCCCC"; }
A.lightgray:active  { text-decoration: none; color: "#CCCCCC"; }
A.lightgray:hover   { text-decoration: underline; color: "#CCCCCC"; }
-->
</style>
<script language="JavaScript">
<!--
function calculateMortgage(){

function IsNum( numstr ) {
	 var downPaymentResult = document.getElementById("downPaymentResult");
	 var mortgageAmountResult = document.getElementById("mortgageAmountResult");
	 var paymentAmount = document.getElementById("paymentAmount");
	 var totalInterest = document.getElementById("totalInterest");

	// Return immediately if an invalid value was passed in
	if (numstr+"" == "undefined" || numstr+"" == "null" || numstr+"" == "")	
		return false;
	var isValid = true;
	var decCount = 0;		
	
	numstr += "";	
	// Loop through string and test each character. If any
	// character is not a number, return a false result.
 	// Include special cases for negative numbers (first char == '-')
	// and a single decimal point (any one char in string == '.').
	for (i = 0; i < numstr.length; i++) {
		// track number of decimal points
		if (numstr.charAt(i) == ".")
			decCount++;
    	if (!((numstr.charAt(i) >= "0") && (numstr.charAt(i) <= "9") ||
				(numstr.charAt(i) == "-") || (numstr.charAt(i) == "."))) {
       	isValid = false;
       	break;
		} else if ((numstr.charAt(i) == "-" && i != 0) ||
				(numstr.charAt(i) == "." && numstr.length == 1) ||
			  (numstr.charAt(i) == "." && decCount > 1)) {
       	isValid = false;
       	break;
      }         	         	
//if (!((numstr.charAt(i) >= "0") && (numstr.charAt(i) <= "9")) ||
   } // END for
   	return isValid;
}  
// end IsNum

// begin calculateMortgage

newP = new String(document.calculate.mortgageamount.value)
mort = Number(document.calculate.mortgageamount.value)


commaExp = /,/gi;
newString = new String ("")
P = newP.replace(commaExp, newString)

if (IsNum(P)){

downPay = ((document.calculate.downPay.value)/100)*(P)
P = (P)-(downPay)

DP = downPay 
DP = DP.toString().replace(/\$|\,/g,'');
if(isNaN(DP))
DP = "0";
sign = (DP == (DP = Math.abs(DP)));
DP = Math.floor(DP*100+0.50000000001);
cents = DP%100;
DP = Math.floor(DP/100).toString();
if(cents<10)
cents = "0" + cents;
for (var i = 0; i < Math.floor((DP.length-(1+i))/3); i++)
DP = DP.substring(0,DP.length-(4*i+3))+','+
DP.substring(DP.length-(4*i+3));

downPaymentResult.firstChild.nodeValue =  (((sign)?'':'-') + '$' + DP + '.' + cents);

MP = (P)
MP = MP.toString().replace(/\$|\,/g,'');
if(isNaN(MP))
MP = "0";
sign = (MP == (MP = Math.abs(MP)));
MP = Math.floor(MP*100+0.50000000001);
cents = MP%100;
MP = Math.floor(MP/100).toString();
if(cents<10)
cents = "0" + cents;
for (var i = 0; i < Math.floor((MP.length-(1+i))/3); i++)
MP = MP.substring(0,MP.length-(4*i+3))+','+
MP.substring(MP.length-(4*i+3));

mortgageAmountResult.firstChild.nodeValue = (((sign)?'':'-') + '$' + MP + '.' + cents);


newI = new String(document.calculate.interestrate.value)
rExp = /%/gi;
newString = new String ("")
I = newI.replace(rExp, newString)

L = document.calculate.mortgagelength.value
J = (I)/(12*100) 
N = (L*12)*(-1) 
M = (P*((J)/(1-(Math.pow(1+(J),(N))))))

newM = (M)
M = M.toString().replace(/\$|\,/g,'');
if(isNaN(M))
M = "0";
sign = (M == (M = Math.abs(M)));
M = Math.floor(M*100+0.50000000001);
cents = M%100;
M = Math.floor(M/100).toString();
if(cents<10)
cents = "0" + cents;
for (var i = 0; i < Math.floor((M.length-(1+i))/3); i++)
M = M.substring(0,M.length-(4*i+3))+','+
M.substring(M.length-(4*i+3));


paymentAmount.firstChild.nodeValue  =  (((sign)?'':'-') + '$' + M + '.' + cents);

T = (newM)*(L)*(12)-(P);
T = T.toString().replace(/\$|\,/g,'');
if(isNaN(T))
T = "0";
sign = (T == (T = Math.abs(T)));
T = Math.floor(T*100+0.50000000001);
cents = T%100;
T = Math.floor(T/100).toString();
if(cents<10)
cents = "0" + cents;
for (var i = 0; i < Math.floor((T.length-(1+i))/3); i++)
T = T.substring(0,T.length-(4*i+3))+','+
T.substring(T.length-(4*i+3));

totalInterest.firstChild.nodeValue =  (((sign)?'':'-') + '$' + T + '.' + cents);

}else{
alert("The Mortgage Amount Field cannot contain special characters ($,@,%,*,etc.).");
}
}


//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="1" bgcolor="#003399"></td>
  </tr>
  <tr bgcolor="#003366"> 
    <td height="54" bgcolor="#FFFFFF"><img src="../images/cstles_logo.gif" width="151" height="91"></td>
  </tr>
  <tr bgcolor="#999999">
    <td class="CastlesTextWhite" height="4" bgcolor="#FFFFFF"><img src="../images/clear10pixel.gif" width="10" height="4"></td>
  </tr>
  <tr bgcolor="#999999"> 
    <td class="CastlesTextWhite" height="20" bgcolor="#4C4691">&nbsp;&nbsp;<b>Mortgage 
      Caculator</b></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td bgcolor="#F8F8F8"> 
      <table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr valign="top"> 
          <td> 
            <form name="calculate" method="post" action="default.asp">
              <table width="346" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">&nbsp;</td>
                  <td width="143" class="CastlesTextBody">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Price of Property:</td>
                  <td width="143" class="CastlesTextBody"> 
                    <input type="text" name="mortgageamount" class="CastlesTextBody" maxlength="20" size="15" value="<%=ListPrice%>">
                  </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Down Payment:</td>
                  <td width="143" class="CastlesTextBody"> 
                    <input type="text" name="downPay" class="CastlesTextBody" maxlength="2" size="3">
                    % </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Interest Rate:</td>
                  <td width="143" class="CastlesTextBody"> 
                    <input type="text" name="interestrate" class="CastlesTextBody" maxlength="5" size="3">
                    % </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Mortgage Length:</td>
                  <td width="143" class="CastlesTextBody"> 
                    <select name="mortgagelength" class="CastlesTextBody">
                      <option value="30" selected>30 years</option>
                      <option value="25">25 years</option>
                      <option value="20">20 years</option>
                      <option value="15">15 years</option>
                      <option value="10">10 years</option>
                      <option value="5">5 years</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right" height="8"></td>
                  <td width="143" class="CastlesTextBody" height="8"></td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">&nbsp;</td>
                  <td width="143" class="CastlesTextBody">&nbsp;<a href="javascript:OnClick=calculateMortgage();" class="gray"><b>Calculate Your 
                    Payment</b> </a></td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right" height="8"></td>
                  <td width="143" class="CastlesTextBody" height="8"></td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Down Payment:</td>
                  <td width="143" class="CastlesTextBody"> 
                    <p id="downPaymentResult">???</p>
                  </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Mortgage Amount: 
                  </td>
                  <td width="143" class="CastlesTextBody"> 
                    <p id="mortgageAmountResult">???</p>
                  </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Monthly Mortgage 
                    Payment:</td>
                  <td width="143" class="CastlesTextBody"> 
                    <p id="paymentAmount">???</p>
                  </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">Total Interest: 
                  </td>
                  <td width="143" class="CastlesTextBody"> 
                    <p id="totalInterest">???</p>
                  </td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">&nbsp;</td>
                  <td width="143" class="CastlesTextBody">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="181" class="CastlesTextBody" align="right">&nbsp;</td>
                  <td width="143" class="CastlesTextBody">&nbsp;</td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
      </table>    </td>
  </tr>
  <tr> 
    <td class="CastlesTextBody"> 
      <table width="300%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#999999">
          <td class="CastlesTextBody" height="20" bgcolor="#4C4691">&nbsp;</td>
        </tr>
        <tr bgcolor="#999999"> 
          <td class="CastlesTextBody" height="4" bgcolor="#4C4691"><img src="../images/clear10pixel.gif" width="10" height="4"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> </td>
  </tr>
</table>
</body>
</html>
