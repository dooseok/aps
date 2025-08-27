<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim RS1
dim RS2
dim SQL
dim CNT1
dim CNT2
dim CNT3

dim s_Process
dim s_Date
dim arrDate(0)

dim MPD_Qty

dim strWidth

dim arrInputSelectG_1
dim arrInputSelect_1
dim arrInputSelectG_2
dim arrInputSelect_2

dim FromTime
dim ToTime

dim BOM_Sub_BS_D_No
dim MPD_Qty_Sum
dim PR_Amount
dim Diff_Sum

dim bgMonitor

dim strBOM_Sub_BS_D_No

s_Process = Request("s_Process")
if s_Process = "" then
	s_Process = "IMD"
end if

s_Date = Request("s_Date")
if s_Date = "" then
	s_Date = date()
end if

arrDate(0) = s_Date

if s_Process="IMD" or s_Process="SMD" then
	arrInputSelectG_2	= split(replace(BasicDataFullTimeStr,"slt>",""),";")
else
	arrInputSelectG_2	= split(replace(BasicDataHalfTimeStr,"slt>",""),";")
end if

select case s_Process
	case "IMD"
		arrInputSelectG_1	= split(replace(BasicDataIMDLine,"slt>",""),";")	
	case "SMD"
		arrInputSelectG_1	= split(replace(BasicDataSMDLine,"slt>",""),";")	
	case "MAN"
		arrInputSelectG_1	= split(replace(BasicDataMANLine,"slt>",""),";")	
	case "ASM"
		arrInputSelectG_1	= split(replace(BasicDataASMLine,"slt>",""),";")
end select

select case s_Process
	case "IMD"
		bgMonitor = "#F2D4D4"
	case "SMD"
		bgMonitor = "#C2E3C6"
	case "MAN"
		bgMonitor = "#C6EBFE"
	case "ASM"
		bgMonitor = "#EADAF7"		
end select

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")
%>
<script language="javascript">
function frmDate_Search_Check()
{
	Show_Progress();
	frmDate_Search.submit();
}

function frmDate_Search_Move(strDate)
{
	frmDate_Search.s_Date.value = strDate;
	Show_Progress();
	frmDate_Search.submit();
}
</script>

<table border=0 cellspacing=1 cellpadding=0 width=420px align=center border=0 bgcolor="#999999">
<form name="frmDate_Search" action="pr_monitor.asp" method="post">
<tr height=25px>
	<td bgcolor=white>
		<table border=0 cellspacing=2 cellpadding=0 width=100% bgcolor="#ffffff">
		<tr>
			<td width=5px>&nbsp;</td>
			<td width=30px align=right>기간</td>
			<td width=5px>&nbsp;</td>
			<td width=45px>
				<%=Make_S_BTN("이전","javascript:frmDate_Search_Move('"&dateadd("d",-1,s_Date)&"');","")%>
			</td>
			<td width=80px>
				<input type="text" name="s_Date" size=10 class="input" readonly value="<%=s_Date%>" onclick="Calendar_D(document.frmDate_Search.s_Date);">
			</td>
			<td width=45px>
				<%=Make_S_BTN("다음","javascript:frmDate_Search_Move('"&dateadd("d",1,s_Date)&"');","")%>
			</td>
			<td width=20px></td>
			<td width=70px align=right>공정</td>
			<td width=50px align=left>
				<select name="s_Process">
				<option value=""<%if s_Process="" then%> selected<%end if%>>-선택-</option>
				<option value="IMD"<%if s_Process="IMD" then%> selected<%end if%>>IMD</option>
				<option value="SMD"<%if s_Process="SMD" then%> selected<%end if%>>SMD</option>
				<option value="MAN"<%if s_Process="MAN" then%> selected<%end if%>>MAN</option>
				<option value="ASM"<%if s_Process="ASM" then%> selected<%end if%>>ASM</option>
				</select>
			</td>
			<td width=50px><%=Make_S_BTN("조회","javascript:frmDate_Search_Check();","")%></td>
			<td width=5px>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<br>
<table width="<%=70+(250*ubound(arrInputSelectG_1))%>px" cellpadding=0 cellspacing=1 border=0 bgcolor="#999999">
<tr bgcolor="<%=bgMonitor%>">
	<td width=70px rowspan=3>&nbsp;</td>
<%
for CNT1 = 0 to ubound(arrDate)
%>
	<td width="<%=(250*ubound(arrInputSelectG_1))%>px" colspan=20>
<%
	response.write arrDate(CNT1)
	
	select case weekday(arrDate(CNT1))
	case 1
		response.write "(일)"
	case 2
		response.write "(월)"
	case 3
		response.write "(화)"
	case 4
		response.write "(수)"
	case 5
		response.write "(목)"
	case 6
		response.write "(금)"
	case 7
		response.write "(토)"
	end select
%>
	</td>
<%
next
%>
</tr>
<tr bgcolor="#eeeeee">
<%
for CNT2 = 0 to ubound(arrInputSelectG_1)
	arrInputSelect_1 = split(arrInputSelectG_1(CNT2),":")
%>
	<td width=250px><%=arrInputSelect_1(0)%></td>
<%
next
%>
</tr>
<tr bgcolor="#eeeeee" height=18px>
<%
for CNT2 = 0 to ubound(arrInputSelectG_1)
	arrInputSelect_1 = split(arrInputSelectG_1(CNT2),":")
%>
	<td width=250px>
		<table width=100% cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td align=center>파트넘버</td>
			<!--<td width=35px align=center>계획</td>
			<td width=35px align=center>실적</td>
			<td width=35px align=center>차이</td>-->
			<td width=70px align=center>생산량</td>
			<td width=5px align=center></td>
		</tr>
		</table>
	</td>
<%
next
%>
</tr>
<%
for CNT1 = 0 to ubound(arrInputSelectG_2)
	if CNT1 > 0 then
		arrInputSelect_2 = split(arrInputSelectG_2(CNT1-1),":")
		FromTime	= replace(right(arrInputSelect_2(1),5),"|","")
	else
		FromTime = "0820"
	end if
	
	arrInputSelect_2 = split(arrInputSelectG_2(CNT1),":")
	ToTime		= replace(right(arrInputSelect_2(1),5),"|","")
	
	ToTime		= left(ToTime,2) & int(right(ToTime,2)) - 1
	
	FromTime	= left(FromTime,2) * 60 + right(FromTime,2) - 500
	ToTime		= left(ToTime,2) * 60 + right(ToTime,2) - 500
%>
<tr bgcolor=white height=60px>
	<td valign=middle bgcolor="#eeeeee"><%=replace(arrInputSelect_2(1),"|",":")%></td>
<%
for CNT2 = 0 to ubound(arrInputSelectG_1)
	arrInputSelect_1 = split(arrInputSelectG_1(CNT2),":")
%>
	<td valign=top>
		<table width=100% cellpadding=0 cellspacing=0 border=0>
<%
	SQL =		"select "&vbcrlf
	SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
	SQL = SQL & "	MPD_Qty_Sum = sum(MPD_Qty) "&vbcrlf
	SQL = SQL & "from "&vbcrlf
	SQL = SQL & "	tbMSE_Plan_Date "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	MPD_Date = '"&s_Date&"' and "&vbcrlf
	SQL = SQL & "	MPD_Process = '"&s_Process&"' and "&vbcrlf
	SQL = SQL & "	MPD_Time = '"&arrInputSelect_2(0)&"' and "&vbcrlf
	SQL = SQL & "	MPD_Line = '"&arrInputSelect_1(0)&"' "&vbcrlf
	SQL = SQL & "group by "&vbcrlf
	SQL = SQL & "	BOM_Sub_BS_D_No "&vbcrlf
	RS1.Open SQL,sys_DBCon
	
	strBOM_Sub_BS_D_No = "'"
	do until RS1.Eof
		BOM_Sub_BS_D_No		= RS1("BOM_Sub_BS_D_No")
		MPD_Qty_Sum			= RS1("MPD_Qty_Sum")
		strBOM_Sub_BS_D_No	= strBOM_Sub_BS_D_No & BOM_Sub_BS_D_No & "','"		
		
		SQL = 		"select "&vbcrlf
		SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
		SQL = SQL & "	PR_Amount "&vbcrlf
		SQL = SQL & "from "&vbcrlf
		SQL = SQL & "	tbProcess_Record "&vbcrlf
		SQL = SQL & "where "&vbcrlf
		SQL = SQL & "	PR_WorkType = '작업' and "&vbcrlf
		SQL = SQL & "	PR_Work_Date = '"&s_Date&"' and "&vbcrlf
		SQL = SQL & "	PR_Process = '"&s_Process&"' and "&vbcrlf
		SQL = SQL & "	PR_Start_Time between '"&FromTime&"' and '"&ToTime&"' and "&vbcrlf
		SQL = SQL & "	PR_Line = '"&arrInputSelect_1(0)&"' and "&vbcrlf
		SQL = SQL & "	BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' "&vbcrlf
		SQL = SQL & "order by "&vbcrlf
		SQL = SQL & "	PR_Start_Time "&vbcrlf
		RS2.Open SQL,sys_DBCon
		if RS2.Eof or RS2.Bof then
			PR_Amount = 0
		else
			PR_Amount = RS2("PR_Amount")
		end if
		RS2.Close
		
		Diff_Sum = int(PR_Amount) - int(MPD_Qty_Sum)
%>
		<tr>
			<td align=center><%=BOM_Sub_BS_D_No%></td>
			<!--<td width=35px align=center><%=MPD_Qty_Sum%></td>-->
			<td width=70px align=center><%=PR_Amount%></td>
			<!--<td width=35px align=center><%=Diff_Sum%></td>-->
			<td width=5px align=center></td>
		</tr>
<%
		RS1.MoveNext
	loop
	RS1.Close
	
	SQL = 		"select "&vbcrlf
	SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
	SQL = SQL & "	PR_Amount "&vbcrlf
	SQL = SQL & "from "&vbcrlf
	SQL = SQL & "	tbProcess_Record "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	PR_Work_Date = '"&s_Date&"' and "&vbcrlf
	SQL = SQL & "	PR_Process = '"&s_Process&"' and "&vbcrlf
	SQL = SQL & "	convert(int,PR_Start_Time) between "&FromTime&" and "&ToTime&" and "&vbcrlf
	SQL = SQL & "	PR_Line = '"&arrInputSelect_1(0)&"' "&vbcrlf
	if len(strBOM_Sub_BS_D_No) > 2 then
		SQL = SQL & "and (BOM_Sub_BS_D_No) not in ("&left(strBOM_Sub_BS_D_No,len(strBOM_Sub_BS_D_No)-2)&") "&vbcrlf
	end if
	SQL = SQL & "order by "&vbcrlf
	SQL = SQL & "	PR_Start_Time "&vbcrlf
	RS2.Open SQL,sys_DBCon
	
	do until RS2.Eof
		BOM_Sub_BS_D_No		= RS2("BOM_Sub_BS_D_No")
		PR_Amount			= RS2("PR_Amount")
%>
		<tr>
			<td align=center><%=BOM_Sub_BS_D_No%></td>
			<!--<td width=35px align=center>&nbsp;</td>-->
			<td width=70px align=center><%=PR_Amount%></td>
			<!--<td width=35px align=center><%=PR_Amount%></td>-->
			<td width=5px align=center></td>
		</tr>
<%
		RS2.MoveNext
	loop
	RS2.Close
%>
		<tr><td colspan=4></td></tr>
		</table>
	</td>
<%
next
%>
</tr>
<%
next
%>
</table>
<%
set RS1 = nothing
set RS2 = nothing
%>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->





