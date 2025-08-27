<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->



<%
dim CNT1
dim CNT2

dim PB_Count

dim RS1
dim SQL
dim arrMSE_Plan(500,2)
dim arrMSE_Buffer(500,2)

arrMSE_Plan(0,0) = "1"
arrMSE_Plan(0,1) = "6871A20774G"
arrMSE_Plan(0,2) = "100"

arrMSE_Plan(1,0) = "1"
arrMSE_Plan(1,1) = "EBR39187712"
arrMSE_Plan(1,2) = "50"

arrMSE_Plan(2,0) = "1"
arrMSE_Plan(2,1) = "6871A20774G"
arrMSE_Plan(2,2) = "100"

arrMSE_Plan(3,0) = "1"
arrMSE_Plan(3,1) = "EBR39187712"
arrMSE_Plan(3,2) = "80"

arrMSE_Plan(4,0) = "1"
arrMSE_Plan(4,1) = "6871A20774G"
arrMSE_Plan(4,2) = "120"

arrMSE_Plan(5,0) = "1" 
arrMSE_Plan(5,1) = "6871A10056D"
arrMSE_Plan(5,2) = "120"

arrMSE_Plan(6,0) = "1"
arrMSE_Plan(6,1) = "EBR44169607"
arrMSE_Plan(6,2) = "70"

arrMSE_Plan(7,0) = "1"
arrMSE_Plan(7,1) = "6871A10056D"
arrMSE_Plan(7,2) = "50"

arrMSE_Plan(8,0) = "1"
arrMSE_Plan(8,1) = "EBR44169607"
arrMSE_Plan(8,2) = "100"

arrMSE_Plan(9,0) = "1"
arrMSE_Plan(9,1) = "6871A10056D"
arrMSE_Plan(9,2) = "500"

dim strPB_PartNo
dim strPB_Count
dim arrPB_PartNo
dim arrPB_Count

set RS1 = server.CreateObject("ADODB.RecordSet")
%>

<%
response.write date() & " MSE 계획"
%>
<table>
<tr>
	<td colspan=3>1라인
	</td>
</tr>
<tr>
	<td>모델</td>
	<td>계획</td>
	<td>실적</td>
</tr>
<%
SQL = "select PB_PartNo, sumPB_Count = sum(PB_Count) from tbPWS_Boxing where PB_Boxing_Date = '"&date()&"' group by PB_PartNo"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strPB_PartNo	= strPB_PartNo	& RS1("PB_PartNo")		& "|"
	strPB_Count		= strPB_Count	& RS1("sumPB_Count")	& "|"
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing

arrPB_PartNo	= split(strPB_PartNo,	"|")
arrPB_Count		= split(strPB_Count,	"|")


for CNT1 = 0 to 9
	PB_Count = 0
	for CNT2 = 0 to ubound(arrPB_PartNo)-1
		if arrMSE_Plan(CNT1,1) = arrPB_PartNo(CNT2) then
			response.write strPB_Count(CNT1,1) & "___" & arrMSE_Plan(CNT1,2)
			if cint(strPB_Count(CNT1,1)) > cint(arrMSE_Plan(CNT1,2)) then
				PB_Count = arrMSE_Plan(CNT1,2)
				strPB_Count(CNT1,1) = strPB_Count(CNT1,1) - arrMSE_Plan(CNT1,2)
			else
				PB_Count = strPB_Count(CNT1,1)
			end if
		end if
	next
%>
<tr>
	<td><%=arrMSE_Plan(CNT1,1)%></td>
	<td><%=arrMSE_Plan(CNT1,2)%></td>
	<td><%=PB_Count%></td>
<tr>
<%
next
set RS1 = nothing
%>
</table>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->





