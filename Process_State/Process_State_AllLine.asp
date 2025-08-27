<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim RS1
dim SQL

'반복문에 사용하기 위한 변수 선언
dim CNT1
dim CNT2
dim CNT3

dim arrNow
dim nNow
dim nSumPlan

dim nAccTime
dim nAccCount

dim PSP_ST
dim PSP_Count

arrNow = split(mid(NOW(),15,10),":")
if instr(NOW(),"오후") > 0 then
	arrNow(0) = arrNow(0) + 12
end if
nNow = arrNow(0) * 3600 + arrNow(1) * 60 + arrNow(2)

set RS1 = Server.CreateObject("ADODB.RecordSet")
%>

<html>
<head>
</head>
<body topmargin=0 leftmargin=0>

<script language="javascript">
</script>

<table width=670px cellpadding=0 cellspacing=1 bgcolor="black">
<tr bgcolor=white>
	<td>라인</td>
	<td>현재생산모델</td>
	<td>계획</td>
	<td>목표</td>
	<td>실적</td>
	<td>달성율</td>
</tr>
<tr bgcolor=white>
	<td>1</td>
	<td>
<%
SQL = "select PRD_PartNo from tbPWS_Raw_Data where PRD_Line='pwsbox1' and PRD_Box_Date = '"&request("s_Work_Date")&"' order by PRD_BOX_Time desc"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	response.write "-"
else
	response.write RS1("PRD_PartNo")
end if
RS1.Close
%>
	</td>
	<td>
<%
SQL = "select sum(PSP_Count) from tbProcess_State_Plan where PSP_Line='1' and PSP_Work_Date = '"&request("s_Work_Date")&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	response.write "-"
else
	response.write RS1(0)
end if
RS1.Close
%>
	</td>
	<td>
<%
nAccTime	= 30000
nAccCount	= 0
SQL = "select PSP_ST,PSP_Count from tbProcess_State_Plan where PSP_Line='1' and PSP_Work_Date = '"&request("s_Work_Date")&"'"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	PSP_ST		= RS1("PSP_ST")	
	PSP_Count	= RS1("PSP_Count")
	
	if nAccTime <= nNow then
		for CNT1 = 1 to PSP_Count
			if nAccTime <= nNow then
				nAccTime  = nAccTime + PSP_ST
				nAccCount = nAccCount + 1
			end if
		next
	end if

	RS1.MoveNext
loop
RS1.Close
response.write nAccCount & "<br>"
%>
	</td>
	<td>
<%
SQL = "select count(PRD_Code) from tbPWS_Raw_Data where PRD_Box_Date is not null and PRD_Line='pwsbox1' and PRD_Box_Date = '"&request("s_Work_Date")&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	nSumPlan = 0
else
	nSumPlan = RS1(0)
end if
RS1.Close
response.write nSumPlan
%>
	</td>
<%
if nAccCount = 0 then
%>
	<td>-</td>
<%
else
%>
	<td><%=formatNumber(nSumPlan * 100 / nAccCount,1)%>%</td>
<%
end if
%>
</tr>
</table>
</body>
</html>

<%
set RS1 = Nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
	