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

dim BOM_Model_BM_D_No

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
	arrInputSelectG_1	= split(replace(BasicDataFullTime,"slt>",""),";")
else
	arrInputSelectG_1	= split(replace(BasicDataHalfTime,"slt>",""),";")
end if

select case s_Process
	case "IMD"
		arrInputSelectG_2	= split(replace(BasicDataIMDLine,"slt>",""),";")	
	case "SMD"
		arrInputSelectG_2	= split(replace(BasicDataSMDLine,"slt>",""),";")	
	case "MAN"
		arrInputSelectG_2	= split(replace(BasicDataMANLine,"slt>",""),";")	
	case "ASM"
		arrInputSelectG_2	= split(replace(BasicDataASMLine,"slt>",""),";")
end select

if instr("-IMD-SMD-","-"&s_Process&"-") > 0 then
	strWidth = "30"
else
	strWidth = "60"
end if

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
<table border=0 cellspacing=1 cellpadding=0 width=420px bgcolor="#999999" align=center>
<form name="frmDate_Search" action="lp_view_table.asp" method="post">
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
			<td width=70px align=right>수정</td>
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
<table width="<%=50+(550*(ubound(arrDate)+1))%>px" cellpadding=0 cellspacing=1 bgcolor="#999999">
<tr bgcolor=white>
	<td width=50px rowspan=2>Line</td>
<%
for CNT1 = 0 to ubound(arrDate)
%>
	<td width=550px>
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
<tr bgcolor=white class="LGE_Plan">
<%
for CNT1 = 0 to ubound(arrDate)
%>
	<td width=550px>
		<table width=100% cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td width=100px>도번</td>
			<td width=150px>공수</td>
<%
	for CNT2 = 0 to ubound(arrInputSelectG_1)
		arrInputSelect_1 = split(arrInputSelectG_1(CNT2),":")
%>
			<td width="<%=strWidth%>px"><%=arrInputSelect_1(1)%></td>
<%
	next
%>
		</tr>
		</table>
	</td>
<%
next
%>
</tr>
<%
for CNT1 = 0 to ubound(arrInputSelectG_2)
	arrInputSelect_2 = split(arrInputSelectG_2(CNT1),":")
%>
<tr bgcolor=white height=19px>
	<td><%=arrInputSelect_2(1)%></td>
<%
	for CNT2 = 0 to ubound(arrDate)	
%>
	<td valign=top>
		<table width=100% cellpadding=0 cellspacing=0 border=0 bgcolor="#999999">	
<%	
		SQL =		"select "&vbcrlf
		SQL = SQL & "	distinct BOM_Model_BM_D_No "&vbcrlf
		SQL = SQL & "from "&vbcrlf
		SQL = SQL & "	tbMSE_Plan_Date "&vbcrlf
		SQL = SQL & "where "&vbcrlf
		SQL = SQL & "	MPD_Process	= '"&s_Process&"' and "&vbcrlf
		SQL = SQL & "	MPD_Date	= '"&arrDate(CNT2)&"' and "&vbcrlf
		SQL = SQL & "	MPD_Line	= '"&arrInputSelect_2(0)&"' "&vbcrlf
		SQL = SQL & "order by BOM_Model_BM_D_No "&vbcrlf
		
		RS1.Open SQL,sys_DBCon
		
		if RS1.Eof or RS1.Bof then
%>
		<tr bgcolor=white><td>&nbsp;</td></tr>
<%	
		else
			do until RS1.Eof
				BOM_Model_BM_D_No = RS1("BOM_Model_BM_D_No")
%>
		<tr bgcolor=white>
			<td width=100px><%=BOM_Model_BM_D_No%></td>
			<td>시간당 100개</td>
<%

				for CNT3 = 0 to ubound(arrInputSelectG_1)
					arrInputSelect_1 = split(arrInputSelectG_1(CNT3),":")
					SQL = 		"select MPD_Qty from tbMSE_Plan_Date "&vbcrlf
					SQL = SQL & "where BOM_Model_BM_D_No = '"&BOM_Model_BM_D_No&"' and MPD_Process='"&s_Process&"' and MPD_Line='"&arrInputSelect_2(0)&"' and MPD_Date='"&arrDate(CNT2)&"' and MPD_Time='"&arrInputSelect_1(0)&"'"&vbcrlf
					SQL = SQL & " "&vbcrlf
					RS2.Open SQL,sys_DBCon
					
					if RS2.Eof or RS2.Bof then
						MPD_Qty = ""
					else
						MPD_Qty = RS2("MPD_Qty")
					end if
					RS2.Close
%>
			<td width="<%=strWidth%>px">
				<input type="text" value="<%=MPD_Qty%>" style="width:<%=strWidth-1%>px;height:17px;text-align:center">
			</td>
<%
				next
%>
		</tr>
<%
				RS1.MoveNext
			loop
		end if
		RS1.Close
%>
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




