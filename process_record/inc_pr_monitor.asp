<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
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
dim PR_Amount_Sum
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


<table width="<%=25+(120*(ubound(arrInputSelectG_1)+1))%>px" cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff">
<tr>
	<td valign=middle align=center>
<table width="<%=25+(120*(ubound(arrInputSelectG_1)+1))%>px" cellpadding=0 cellspacing=1 border=0 bgcolor="#666666">


<tr bgcolor="<%=bgMonitor%>" height=18px>
	<td width=25px></td>
<%
for CNT2 = 0 to ubound(arrInputSelectG_1)
	arrInputSelect_1 = split(arrInputSelectG_1(CNT2),":")
%>
	<td width=120px style="font-family:arial;font-size:10px"><%=arrInputSelect_1(0)%></td>
<%
next
%>
</tr>
<%
for CNT1 = 0 to ubound(arrInputSelectG_2)
	arrInputSelect_2 = split(arrInputSelectG_2(CNT1),":")
	
	FromTime	= replace(left(arrInputSelect_2(1),5),"|","")
	ToTime		= replace(right(arrInputSelect_2(1),5),"|","")
	
	ToTime		= left(ToTime,2) & int(right(ToTime,2)) - 1
%>
<tr bgcolor="<%=bgMonitor%>" height=50px>
	<td width=25px valign=middle bgcolor="<%=bgMonitor%>" style="font-family:arial;font-size:10px"><%=arrInputSelect_2(0)%></td>
<%
for CNT2 = 0 to ubound(arrInputSelectG_1)
	arrInputSelect_1 = split(arrInputSelectG_1(CNT2),":")
%>
	<td width=120px valign=top>
		<table width=120px cellpadding=0 cellspacing=0 border=0>
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
		SQL = SQL & "	PR_Amount_Sum = sum(PR_Amount) "&vbcrlf
		SQL = SQL & "from "&vbcrlf
		SQL = SQL & "	tbProcess_Record "&vbcrlf
		SQL = SQL & "where "&vbcrlf
		SQL = SQL & "	PR_WorkType = 'ÀÛ¾÷' and "&vbcrlf
		SQL = SQL & "	PR_Work_Date = '"&s_Date&"' and "&vbcrlf
		SQL = SQL & "	PR_Process = '"&s_Process&"' and "&vbcrlf
		SQL = SQL & "	PR_Start_Time between '"&FromTime&"' and '"&ToTime&"' and "&vbcrlf
		SQL = SQL & "	PR_Line = '"&arrInputSelect_1(0)&"' and "&vbcrlf
		SQL = SQL & "	BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' "&vbcrlf
		SQL = SQL & "group by "&vbcrlf
		SQL = SQL & "	BOM_Sub_BS_D_No "&vbcrlf
		RS2.Open SQL,sys_DBCon
		if RS2.Eof or RS2.Bof then
			PR_Amount_Sum = 0
		else
			PR_Amount_Sum = RS2("PR_Amount_Sum")
		end if
		RS2.Close
		
		Diff_Sum = int(PR_Amount_Sum) - int(MPD_Qty_Sum)
%>
		<tr>
			<td align=center style="font-family:arial;font-size:10px"><%=BOM_Sub_BS_D_No%></td>
			<td width=30px align=center style="font-family:arial;font-size:10px"><%=MPD_Qty_Sum%></td>
		</tr>
<%
		RS1.MoveNext
	loop
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
	</td>
</tr>
</table>
<%
set RS1 = nothing
set RS2 = nothing
%>
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->





