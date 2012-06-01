<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/templates/castlessystemcnekt.asp" -->
<%
'setting variables
SubscriptionTypeID = 1
SubscriptionTransactionDateTime = "12/18/2002"
'SubscriptionTransactionDateTime = now

'calculate the first issue and subscription expiration
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_SubscriptionType_PeriodMonths"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.Parameters.Append .CreateParameter("@SubscriptionTypeID", 200, 1,200,SubscriptionTypeID)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set SubscriptionPeriod = .Execute()
		End With
		Set Command1 = Nothing
		
		Set Command1 = Server.CreateObject("ADODB.Command")
		With Command1	
			.ActiveConnection = Connect
			.CommandText = "Castles_System_FirstIssueCalculator"
			.Parameters.Append .CreateParameter("@RETURN_VALUE", 3, 4)
			.CommandType = 4
			.CommandTimeout = 0
			.Prepared = true
			Set FirstIssueCalc = .Execute()
		End With
		Set Command1 = Nothing
		
		SubscriptionPeriodMonths = SubscriptionPeriod.Fields.Item("SubscriptionPeriodMonths").Value
		DateSpan = FirstIssueCalc.Fields.Item("DateSpan").Value
		CutoffDay = FirstIssueCalc.Fields.Item("CutoffDay").Value
		
		CurrentMonth = DatePart("m",SubscriptionTransactionDateTime)
		AfterSpanDate = DateAdd("d",DateSpan,SubscriptionTransactionDateTime)
		AfterSpanYear = DatePart("yyyy",AfterSpanDate)
		AfterSpanMonth = DatePart("m",AfterSpanDate)
		AfterSpanDay = DatePart("d",AfterSpanDate)
		if AfterSpanDay < CutoffDay then
			FirstIssueMonthYear = AfterSpanMonth & "/1/" & AfterSpanYear
		elseif AfterSpanDay >= CutoffDay then
			FirstIssueMonthYear = (AfterSpanMonth+1) & "/1/" & AfterSpanYear
		end if
		ExpirationMonthYear = DateAdd("m",SubscriptionPeriodMonths,FirstIssueMonthYear)
		response.write "now is: " & now & "<BR>"
		response.write "SubscriptionTransactionDateTime is: " & SubscriptionTransactionDateTime & "<BR>"		
		response.write "DateSpan is: " & DateSpan & "<BR>"
		response.write "AfterSpanDate is: " & AfterSpanDate & "<BR>"
		response.write "CutoffDay is: " & CutoffDay & "<BR>"
		response.write "SubscriptionPeriodMonths is: " & SubscriptionPeriodMonths & "<BR>"
		response.write "FirstIssueMonthYear is: " & FirstIssueMonthYear & "<BR>"
		response.write "ExpirationMonthYear is: " & ExpirationMonthYear & "<BR>"
%>
