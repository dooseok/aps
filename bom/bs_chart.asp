<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
dim RS1
dim SQL

dim BS_Code
dim PR_Code

dim BS_D_No
dim arrQtyByDate
dim nDay
dim nDay_To
dim s_Date

BS_Code = request("BS_Code")
PR_Code	= request("PR_Code")
s_Date	= request("s_Date")

set RS1 = Server.CreateObject("ADODB.RecordSet")

if s_Date = "" then
	s_Date = Date()
end if

if BS_Code = "" then
	SQL = "select bom_sub_bs_d_no from tbProcess_Record where PR_Code = '"&PR_Code&"'"
	RS1.Open SQL,sys_DBCon
	BS_D_No = RS1("bom_sub_bs_d_no")
	RS1.Close
else
	SQL = "select bs_d_no from tbBOM_Sub where bs_code = '"&BS_Code&"'"
	RS1.Open SQL,sys_DBCon
	BS_D_No = RS1("BS_D_No")
	RS1.Close
end if

set RS1 = nothing

arrQtyByDate = split(getQtyByDate(BS_D_No,left(s_Date,7)),"|%|")
%>
	
<script language="javascript">
function MoveNextMonth()
{
	location.href="bs_chart.asp?PR_Code=<%=PR_Code%>&BS_Code=<%=BS_Code%>&s_Date=<%=dateadd("m",1,s_Date)%>";
}
function MovePreMonth()
{
	location.href="bs_chart.asp?PR_Code=<%=PR_Code%>&BS_Code=<%=BS_Code%>&s_Date=<%=dateadd("m",-1,s_Date)%>";
}
</script>
<br>
<table width=700px cellpadding=0 cellspacing=1 bgcolor=white>
<tr>
	<td style="font-size:18pt"><%=BS_D_No%> 재공재고 변동현황</td>
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
	<td width=55px align=center bgcolor="#eeeeee">IMD</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#ffffff<%end if%>"><%if datediff("d",date(),left(s_Date,7)&"-"&nDay) > 0 then%>&nbsp;<%else%><%=mid(arrQtyByDate(nDay-1),instr(arrQtyByDate(nDay-1),"I")+1,instr(arrQtyByDate(nDay-1),"S")-instr(arrQtyByDate(nDay-1),"I")-1)%><%end if%></td><%next%>
</tr>
<tr bgcolor=white height=30px>
	<td width=55px align=center bgcolor="#eeeeee">SMT</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#ffffff<%end if%>"><%if datediff("d",date(),left(s_Date,7)&"-"&nDay) > 0 then%>&nbsp;<%else%><%=mid(arrQtyByDate(nDay-1),instr(arrQtyByDate(nDay-1),"S")+1,instr(arrQtyByDate(nDay-1),"M")-instr(arrQtyByDate(nDay-1),"S")-1)%><%end if%></td><%next%>
</tr>
<tr bgcolor=white height=30px>
	<td width=55px align=center bgcolor="#eeeeee">MAN</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#ffffff<%end if%>"><%if datediff("d",date(),left(s_Date,7)&"-"&nDay) > 0 then%>&nbsp;<%else%><%=mid(arrQtyByDate(nDay-1),instr(arrQtyByDate(nDay-1),"M")+1,instr(arrQtyByDate(nDay-1),"A")-instr(arrQtyByDate(nDay-1),"M")-1)%><%end if%></td><%next%>
</tr>
<tr bgcolor=white height=30px>
	<td width=55px align=center bgcolor="#eeeeee">ASM</td><%for nDay = 1 to nDay_To%><td width=35px align=center bgcolor="<%if datediff("d",date(),left(s_Date,7)&"-"&nDay) = 0 then%>yellow<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 7 then%>skyblue<%elseif WeekDay(left(s_Date,7)&"-"&nDay) = 1 then%>pink<%else%>#ffffff<%end if%>"><%if datediff("d",date(),left(s_Date,7)&"-"&nDay) > 0 then%>&nbsp;<%else%><%=mid(arrQtyByDate(nDay-1),instr(arrQtyByDate(nDay-1),"A")+1,len(arrQtyByDate(nDay-1))-instr(arrQtyByDate(nDay-1),"A"))%><%end if%></td><%next%>
</tr>
</table>

<%
function getQtyByDate(BS_D_No, strDate)
	dim strQtyByDate
	dim nDay
	dim nDay_To
	if instr("-01-03-05-07-08-10-12-",mid(s_Date,6,2)) > 0 then
		nDay_To = 31
	elseif instr("-04-06-09-11-",mid(s_Date,6,2)) > 0 then
		nDay_To = 30
	else
		nDay_To = 28
	end if
	strQtyByDate = ""
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	for nDay = 1 to nDay_To
		SQL = "select top 1 bsqh_imd_qty, bsqh_smd_qty, bsqh_man_qty, bsqh_asm_qty from tbBOM_Sub_Qty_History where BSQH_Date = '"&strDate&"-"&nDay&"' and BOM_Sub_BS_D_No = '"&BS_D_No&"' order by bsqh_code desc"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			strQtyByDate = strQtyByDate & "I0S0M0A0|%|"
		else
			strQtyByDate = strQtyByDate & "I"&RS1("bsqh_imd_qty")&"S"&RS1("bsqh_smd_qty")&"M"&RS1("bsqh_man_qty")&"A"&RS1("bsqh_asm_qty")&"|%|"
		end if
		RS1.Close
	next
	set RS1 = nothing
	
	getQtyByDate = strQtyByDate
end function
%>



<!-- include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->