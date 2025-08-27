<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
dim RS1
dim SQL

dim BS_Code
dim s_BS_D_No
dim PR_Code

dim BS_D_No
dim arrCompleteQtyByDate
dim arrDLVQtyByDate
dim arrStockQtyByDate

dim nDay
dim nDay_To
dim s_Date

BS_Code = request("BS_Code")
PR_Code	= request("PR_Code")
s_BS_D_No = request("s_BS_D_No")
s_Date	= request("s_Date")

set RS1 = Server.CreateObject("ADODB.RecordSet")

if s_Date = "" then
	s_Date = Date()
end if

if PR_Code <> "" then
	SQL = "select bom_sub_bs_d_no from tbProcess_Record where PR_Code = '"&PR_Code&"'"
	RS1.Open SQL,sys_DBCon
	BS_D_No = RS1("bom_sub_bs_d_no")
	RS1.Close
elseif s_BS_D_No <> "" then
	BS_D_No = s_BS_D_No
elseif BS_Code <> "" then
	SQL = "select bs_d_no from tbBOM_Sub where bs_code = '"&BS_Code&"'"
	RS1.Open SQL,sys_DBCon
	BS_D_No = RS1("BS_D_No")
	RS1.Close
end if

set RS1 = nothing

arrCompleteQtyByDate = split(getCompleteQtyByDate(BS_D_No,left(s_Date,7)),",")
arrDLVQtyByDate = split(getDLVQtyByDate(BS_D_No,left(s_Date,7)),",")
arrStockQtyByDate = split(getStockQtyByDate(BS_D_No,left(s_Date,7)),",")
%>
	
<script language="javascript">
function MoveNextMonth()
{
	location.href="bs_qty_chart.asp?PR_Code=<%=PR_Code%>&BS_Code=<%=BS_Code%>&s_Date=<%=dateadd("m",1,s_Date)%>";
}
function MovePreMonth()
{
	location.href="bs_qty_chart.asp?PR_Code=<%=PR_Code%>&BS_Code=<%=BS_Code%>&s_Date=<%=dateadd("m",-1,s_Date)%>";
}
</script>
<br>
<table width=700px cellpadding=0 cellspacing=1 bgcolor=white>
<tr>
	<td style="font-size:18pt"><%=BS_D_No%> 재고 변동현황</td>
</tr>
</table>
<br>
<table width=700px cellpadding=0 cellspacing=1 bgcolor=white>
<tr>
	<td>
		<table cellpadding=0 cellspacing=1 bgcolor=white>
		<tr>
			<td width=77px><%=Make_BTN("이전달","MovePreMonth()","")%></td>
			<td width=100px><b><%=replace(left(s_Date,7),"-","년 ")%>월</b></td>
			<td width=77px><%=Make_BTN("다음달","MoveNextMonth()","")%></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<%
if instr("-01-03-05-07-08-10-12-",mid(s_Date,6,2)) > 0 then
	nDay_To = 31
elseif instr("-04-06-09-11-",mid(s_Date,6,2)) > 0 then
	nDay_To = 30
else
	nDay_To = 28
end if
%>
<table width="<%=55+35*nDay_To%>"px cellpadding=0 cellspacing=1 bgcolor=black>
<tr bgcolor=white height=30px>
	<td width=55px align=center bgcolor="#eeeeee">&nbsp;</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#eeeeee<%end if%>"><%=nDay%></td><%next%>
</tr>
<tr bgcolor=white height=30px>
	<td width=55px align=center bgcolor="#eeeeee">생산</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#ffffff<%end if%>"><%if datediff("d",date(),left(s_Date,7)&"-"&nDay) > 0 then%>&nbsp;<%else%><%=arrCompleteQtyByDate(nDay-1)%><%end if%></td><%next%>
</tr>
<tr bgcolor=white height=30px>
	<td width=55px align=center bgcolor="#eeeeee">출고</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#ffffff<%end if%>"><%if datediff("d",date(),left(s_Date,7)&"-"&nDay) > 0 then%>&nbsp;<%else%><%=arrDLVQtyByDate(nDay-1)%><%end if%></td><%next%>
</tr>
<tr bgcolor=white height=30px>
	<td width=55px align=center bgcolor="#eeeeee">재고</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#ffffff<%end if%>"><%if datediff("d",date(),left(s_Date,7)&"-"&nDay) > 0 then%>&nbsp;<%else%><%=arrStockQtyByDate(nDay-1)%><%end if%></td><%next%>
</tr>
</table>

<%
function getCompleteQtyByDate(BS_D_No, strDate)
	dim strCompleteQtyByDate
	dim nDay
	dim nDay_To
	dim strQDate
	if instr("-01-03-05-07-08-10-12-",mid(s_Date,6,2)) > 0 then
		nDay_To = 31
	elseif instr("-04-06-09-11-",mid(s_Date,6,2)) > 0 then
		nDay_To = 30
	else
		nDay_To = 28
	end if
	strCompleteQtyByDate = ""
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	for nDay = 1 to nDay_To
	
		if nDay < 10 then
		strQDate = strDate&"-0"&nDay
		else
		strQDate = strDate&"-"&nDay
		end if
		SQL = "select cntPRD_Barcode = count(PRD_Barcode) from tbPWS_Raw_Data where PRD_Box_Date = '"&strQDate&"' and PRD_PartNO = '"&BS_D_No&"'"
		
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			strCompleteQtyByDate = strCompleteQtyByDate & "0,"
		else
			strCompleteQtyByDate = strCompleteQtyByDate & RS1("cntPRD_Barcode")&","
		end if
		RS1.Close
	next
	set RS1 = nothing
	
	getCompleteQtyByDate = strCompleteQtyByDate
end function
%>

<%
function getDLVQtyByDate(BS_D_No, strDate)
	dim strDLVQtyByDate
	dim nDay
	dim nDay_To
	dim strQDate
	if instr("-01-03-05-07-08-10-12-",mid(s_Date,6,2)) > 0 then
		nDay_To = 31
	elseif instr("-04-06-09-11-",mid(s_Date,6,2)) > 0 then
		nDay_To = 30
	else
		nDay_To = 28
	end if
	strDLVQtyByDate = ""
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	for nDay = 1 to nDay_To
	
		if nDay < 10 then
		strQDate = strDate&"-0"&nDay
		else
		strQDate = strDate&"-"&nDay
		end if
		SQL = "select sumPR_Amount = isnull(sum(PR_Amount),0) from tbProcess_Record where BOM_Sub_BS_D_No='"&BS_D_No&"' and PR_Work_Date = '"&strQDate&"' and PR_Process = 'DLV'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			strDLVQtyByDate = strDLVQtyByDate & "0,"
		else
			strDLVQtyByDate = strDLVQtyByDate & RS1("sumPR_Amount")&","
		end if
		RS1.Close
	next
	set RS1 = nothing
	
	getDLVQtyByDate = strDLVQtyByDate
end function
%>

<%
function getStockQtyByDate(BS_D_No, strDate)
	dim strStockQtyByDate
	dim nDay
	dim nDay_To
	dim strQDate
	if instr("-01-03-05-07-08-10-12-",mid(s_Date,6,2)) > 0 then
		nDay_To = 31
	elseif instr("-04-06-09-11-",mid(s_Date,6,2)) > 0 then
		nDay_To = 30
	else
		nDay_To = 28
	end if
	strStockQtyByDate = ""
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	for nDay = 1 to nDay_To
		if nDay < 10 then
		strQDate = strDate&"-0"&nDay
		else
		strQDate = strDate&"-"&nDay
		end if
		SQL = "select top 1 bsql_qty = bsql_man_qty+bsql_asm_qty from tbBOM_Sub_Qty_Log where BSQL_update_Date <= '"&strQDate&"' and BOM_Sub_BS_D_No = '"&BS_D_No&"' order by BSQL_update_Date desc"
		
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			strStockQtyByDate = strStockQtyByDate & "0,"
		else
			strStockQtyByDate = strStockQtyByDate & RS1("bsql_qty")&","
		end if
		RS1.Close
	next
	set RS1 = nothing
	
	getStockQtyByDate = strStockQtyByDate
end function
%>


<!-- include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->