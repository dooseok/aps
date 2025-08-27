<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim arrLine
dim CNT1
dim s_Print_By
dim s_PR_Work_Date
dim s_PR_Work_Date2
dim s_PR_Work_Date_SQL

s_Print_By = Request("s_Print_By")
s_PR_Work_Date = Request("s_PR_Work_Date")
s_PR_Work_Date2 = Request("s_PR_Work_Date2")
response.write s_PR_Work_Date2
if len(s_PR_Work_Date2) = "22" then
	s_PR_Work_Date_SQL =  " between '"&Left(s_PR_Work_Date2,10)&"' and '"&Right(s_PR_Work_Date2,10)&"'"
	s_PR_Work_Date = replace(s_PR_Work_Date2,", "," - ")
else
	s_PR_Work_Date_SQL = s_PR_Work_Date
end if
%>

<table width=960px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr style="font-weight:bold;">
	<td width=250px>&nbsp;</td>
	<td width=229px align=center style="font-size:25px;">
		<%if s_Print_By = "DLV" then%>영업일보<%else%>생산일보<%end if%>
	</td>
	<td width=250px valign=bottom align=right>
		<div class="PR_Print">
		<table width=100% cellpadding=0 cellspacing=0 border=0 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
		<tr style="font-weight:bold;">
			<td><%if s_Print_By = "DLV" then%>납품일<%else%>제조일<%end if%> : <%=s_PR_Work_Date%><br>&nbsp;출력일 : <%=date()%>&nbsp;<br></td>
		</tr>
		</table>
	</td>
	<td width=191px align=right style="font-size:15px;">
		<table class="pi_print_2" width=185px cellpadding=0 bordercolor="gray" cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
		<tr bgcolor=white>
			<td width=25px rowspan=2>결<br>재</td>
			<td width=40px>팀 장</td>
			<td width=40px>임 원</td>
			<td width=40px>사 장</td>
		</tr>
		<tr bgcolor=white height=40px>
			<td>&nbsp;</td>
			<td>&nbsp;</td>		
			<td>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
if s_Print_By = "SMD" then
	call Common_PR_List_Summary(s_PR_Work_Date_SQL,"SMD")
	arrLine = GetLine(s_PR_Work_Date_SQL,"SMD")
	for CNT1 = 0 to ubound(arrLine)
		if arrLine(CNT1) <> "" then
			call Common_PR_List(s_PR_Work_Date_SQL,"SMD",arrLine(CNT1))
		end if
	next
elseif s_Print_By = "IMD" then
	call Common_PR_List_Summary(s_PR_Work_Date_SQL,"IMD")
	arrLine = GetLine(s_PR_Work_Date_SQL,"IMD")
	for CNT1 = 0 to ubound(arrLine)
		if arrLine(CNT1) <> "" then
			call Common_PR_List(s_PR_Work_Date_SQL,"IMD",arrLine(CNT1))
		end if
	next
elseif s_Print_By = "DLV" then
	call Common_PR_List_Summary(s_PR_Work_Date_SQL,"DLV")
	arrLine = GetLine(s_PR_Work_Date_SQL,"DLV")
	for CNT1 = 0 to ubound(arrLine)
		if arrLine(CNT1) <> "" then
			call Common_PR_List_DLV(s_PR_Work_Date_SQL,"DLV",arrLine(CNT1))
		end if
	next
else
	call Common_PR_List_Summary(s_PR_Work_Date_SQL,"MAN")
	call Common_PR_List_Summary(s_PR_Work_Date_SQL,"ASM")
	arrLine = GetLine(s_PR_Work_Date_SQL,"MAN")
	for CNT1 = 0 to ubound(arrLine)
		if arrLine(CNT1) <> "" then
			call Common_PR_List(s_PR_Work_Date_SQL,"MAN",arrLine(CNT1))
		end if
	next
	arrLine = GetLine(s_PR_Work_Date_SQL,"ASM")
	for CNT1 = 0 to ubound(arrLine)
		if arrLine(CNT1) <> "" then
			call Common_PR_List(s_PR_Work_Date_SQL,"ASM",arrLine(CNT1))
		end if
	next
end if
%>

<%
function GetLine(strPR_Work_Date, PR_Process)
	dim SQL
	dim RS1
	dim strPR_Line
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select distinct PR_Line from tbProcess_Record where PR_Process='"&PR_Process&"' and "
	if instr(strPR_Work_Date,"between") > 0 then
		SQL = SQL & "PR_Work_Date "&strPR_Work_Date&" "
	else
		SQL = SQL & "PR_Work_Date = '"&strPR_Work_Date&"' "
	end if
	SQL = SQL & "order by PR_Line asc "
	'response.write SQL
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof 
		strPR_Line = strPR_Line & RS1("PR_Line") &"|"
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing

	GetLine = split(strPR_Line,"|")
end function
%>

<%
sub Common_PR_List_Summary(strPR_Work_Date, strPR_Process)
	dim RS1
	dim RS2
	
	dim SQL
	
	dim sumPR_Rest_Time
	dim sumPR_Loss_Ctrl
	dim sumPR_Amount
	dim sumPR_Price
	dim sumPR_Calc_Point
	
	dim Neung_Yul
	dim HoeSu_Yul
	dim SilDong_Yul
	dim GaDong_Yul
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
%>
<div class="PR_Print">
<table width=960px cellpadding=0 cellspacing=0 border=0 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr style="font-weight:bold;">
	<td align=left>
		날짜 :
<%
	response.write s_PR_Work_Date
%>
		&nbsp;&nbsp;&nbsp;
		공정 :
<%
	select case strPR_Process
	case "IMD"
		response.write "IMD"
	case "SMD"
		response.write "SMD"
	case "MAN"
		response.write "수삽"
	case "ASM"
		response.write "조립"
	case "DLV"
		response.write "영업"
	end select
%>
	</td>
</tr>
</table>
<%
if strPR_Process = "DLV" then
%>
<table width=1000px cellpadding=0 cellspacing=0 border=0>
<tr>
	<td width=10px></td>
	<td align=left>
		<table width=270px cellpadding=0 cellspacing=0 border=1 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
		<tr style="font-weight:bold;" bgcolor="skyblue" >
			<td width=100px>납품처</td>
			<td width=70px>납품<br>수량</td>
			<td>생산금액<br>(LG판가기준)</td>
		</tr>
		
		<%
				if instr(strPR_Work_Date,"between") > 0 then
					SQL = "select * from vwPR_Report_Analysis where PR_Work_Date "&strPR_Work_Date&" and PR_Process = '"&strPR_Process&"' order by PR_Line asc"
				else
					SQL = "select * from vwPR_Report_Analysis where PR_Work_Date = '"&strPR_Work_Date&"' and PR_Process = '"&strPR_Process&"' order by PR_Line asc"
				end if
				RS1.Open SQL,sys_DBcon
				if RS1.Eof or RS1.Bof then
		%>
		<tr>
			<td colspan=20>등록된 납품 실적이 없습니다.</td>
		</tr>
		<%
				else
					do until RS1.Eof
		%>
		<tr>
			<td><%=RS1("PR_Line")%></td>
			<td align=right><%=RS1("PR_Amount")%>개&nbsp;</td>
			<td align=right><%=customFormatCurrency(RS1("PR_Price"))%>&nbsp;</td>
		</tr>
		<%
						sumPR_Amount		= sumPR_Amount		+ RS1("PR_Amount")
						sumPR_Price			= sumPR_Price		+ RS1("PR_Price")
						RS1.MoveNext
					loop
		%>
		<tr style="font-weight:bold;" bgcolor=pink>
			<td>총계</td>
			<td align=right><%=sumPR_Amount%>개&nbsp;</td>
			<td align=right><%=customFormatCurrency(sumPR_Price)%>&nbsp;</td>
		</tr>
		<%
				end if
				RS1.Close
				set RS1 = nothing
		%>
		</table>
	</td>
</tr>
</table>
<%
else
%>
<table width=960px cellpadding=0 cellspacing=0 border=1 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr style="font-weight:bold;" bgcolor="skyblue" >
	<td width=32px>라인</td>
	<td width=32px>직접<br>인원</td>
	<td width=32px>간접<br>인원</td>
	<td width=42px>시작<br>시각</td>
	<td width=42px>종료<br>시각</td>
	<td width=42px>근무<br>시간</td>
	<td width=42px>무작업<br>시간</td>
	<td width=47px>생산<br>수량</td>
	<td width=47px>직접<br>공수</td>
	<td width=47px>간접<br>공수</td>
	<td width=47px>무작업<br>공수</td>
	<td width=47px>실동<br>공수</td>
	<td width=47px>재작업<br>공수</td>
	<td width=47px>순작업<br>공수</td>
	<td width=47px>회수<br>공수</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td width=55px>능률</td>
	<td width=55px>회수율</td>
<%
		else
%>
	<td width=55px>양품율</td>
	<td width=55px>설비<br>효율</td>
<%
		end if
%>
	<td width=55px>실동율</td>
	<td width=55px>가동율</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td>생산금액<br>(LG판가기준)</td>
<%
		else
%>
	<td>총점수</td>
<%
		end if
%>
</tr>

<%
		if instr(strPR_Work_Date,"between") > 0 then
			SQL = "select * from vwPR_Report_Analysis where PR_Work_Date "&strPR_Work_Date&" and PR_Process = '"&strPR_Process&"' order by PR_Line asc"
		else
			SQL = "select * from vwPR_Report_Analysis where PR_Work_Date = '"&strPR_Work_Date&"' and PR_Process = '"&strPR_Process&"' order by PR_Line asc"
		end if
		RS1.Open SQL,sys_DBcon
		if RS1.Eof or RS1.Bof then
%>
<tr>
	<td colspan=20>등록된 제조 실적이 없습니다.</td>
</tr>
<%
		else
			do until RS1.Eof
				Neung_Yul	= RS1("Neung_Yul")
				if len(Neung_Yul) - instr(Neung_Yul,".") = 1 then
					Neung_Yul = Neung_Yul & "0"
				end if
				HoeSu_Yul	= RS1("HoeSu_Yul")
				if len(HoeSu_Yul) - instr(HoeSu_Yul,".") = 1 then
					HoeSu_Yul = HoeSu_Yul & "0"
				end if
				SilDong_Yul	= RS1("SilDong_Yul")
				if len(SilDong_Yul) - instr(SilDong_Yul,".") = 1 then
					SilDong_Yul = SilDong_Yul & "0"
				end if
				GaDong_Yul	= RS1("GaDong_Yul")
				if len(GaDong_Yul) - instr(GaDong_Yul,".") = 1 then
					GaDong_Yul = GaDong_Yul & "0"
				end if
			
%>
<tr>
	<td><%=RS1("PR_Line")%></td>
	<td align=right><%=RS1("PR_Worker")%>명&nbsp;</td>
	<td align=right><%=RS1("PR_Supporter")%>명&nbsp;</td>
	<td><%=RS1("PR_Start_Time")%></td>
	<td><%=RS1("PR_End_Time")%></td>
	<td><%=RS1("PR_Diff_Time")%></td>
	<td><%=RS1("PR_Loss_Ctrl")%></td>
	<td align=right><%=RS1("PR_Amount")%>개&nbsp;</td>
	<td align=right><%=RS1("JikJup_GongSu")%>&nbsp;</td>
	<td align=right><%=RS1("Ganjup_GongSu")%>&nbsp;</td>
	<td align=right><%=RS1("MuJakUp_GongSu")%>&nbsp;</td>
	<td align=right><%=RS1("SilDong_GongSu")%>&nbsp;</td>
	<td align=right><%=RS1("JeJakUp_GongSu")%>&nbsp;</td>
	<td align=right><%=RS1("SunJakUp_GongSu")%>&nbsp;</td>
	<td align=right><%=round(RS1("HoeSu_GongSu"))%>&nbsp;</td>
	<td align=right><%=Neung_Yul%>%&nbsp;</td>
	<td align=right><%=HoeSu_Yul%>%&nbsp;</td>
	<td align=right><%=SilDong_Yul%>%&nbsp;</td>
	<td align=right><%=GaDong_Yul%>%&nbsp;</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td align=right><%=customFormatCurrency(RS1("PR_Price"))%>&nbsp;</td>
<%
		elseif instr("-IMD-","-"&strPR_Process&"-") > 0 then
%>
	<td align=right>&nbsp;</td>
<%
		else
%>
	<td align=right><%=CustomFormatComma(RS1("PR_Calc_Point"))%>점&nbsp;</td>
<%
		end if
%>
</tr>
<%
				sumPR_Loss_Ctrl		= sumPR_Loss_Ctrl	+ int(left(RS1("PR_Loss_Ctrl"),2))*60 + int(right(RS1("PR_Loss_Ctrl"),2))
				sumPR_Amount		= sumPR_Amount		+ RS1("PR_Amount")
				sumPR_Price			= sumPR_Price		+ RS1("PR_Price")
				sumPR_Calc_Point	= sumPR_Calc_Point	+ RS1("PR_Calc_Point")
		
				RS1.MoveNext
			loop
			sumPR_Loss_Ctrl = right("0" & int(sumPR_Loss_Ctrl / 60),2) & ":" & right("0" & int(sumPR_Loss_Ctrl mod 60),2)
%>
<tr style="font-weight:bold;" bgcolor=pink>
	<td>총계</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td><%=sumPR_Loss_Ctrl%></td>
	<td align=right><%=sumPR_Amount%>개&nbsp;</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
<%
			if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>	
	<td align=right><%=customFormatCurrency(sumPR_Price)%>&nbsp;</td>
<%
			elseif instr("-IMD-","-"&strPR_Process&"-") > 0 then
%>
	<td align=right>&nbsp;</td>
<%
			else
%>
	<td align=right><%=CustomFormatComma(sumPR_Calc_Point)%>점&nbsp;</td>
<%
			end if
%>
</tr>
<%
		end if
		RS1.Close
		set RS1 = nothing
%>
</table>
<%
end if
%>
</div>
<br><br>
<%
end sub
%>

<%
sub Common_PR_List(strPR_Work_Date, strPR_Process, strPR_Line)
	
	dim SQL
	dim RS1
	dim RS2
	dim CNT1
	dim halfCNT1
	
	dim JikJup_GongSu
	dim GanJup_GongSu
	dim HoeSu_GongSu
	
	dim halfJikJup_GongSu
	dim halfGanJup_GongSu
	dim halfHoeSu_GongSu
	
	dim Diff_Of_maxPR_End_Time_And_minPR_Start_Time
	dim halfDiff_Of_maxPR_End_Time_And_minPR_Start_Time
		
	dim minPR_Start_Time
	dim maxPR_End_Time
	
	dim halfminPR_Start_Time
	dim halfmaxPR_End_Time
	
	dim sumPR_Worker_CNT_Time
	dim sumPR_Supporter_CNT_Time
	dim sumPR_Amount_ST
	dim sumPR_Amount_Point
	
	dim sumPR_Amount
	dim sumPR_Worker_CNT
	dim sumPR_Supporter_CNT
	dim sumPR_Loss_Time
	dim sumPR_Time_Diff
	dim sumPR_Time_Diff_Expected
	dim sumPR_Calc_Point
	dim sumPR_Hoesu_Yul
	dim sumPR_Price
	
	dim halfsumPR_Worker_CNT_Time
	dim halfsumPR_Supporter_CNT_Time
	dim halfsumPR_Amount_ST
	dim halfsumPR_Amount_Point
	
	dim halfsumPR_Amount
	dim halfsumPR_Worker_CNT
	dim halfsumPR_Supporter_CNT
	dim halfsumPR_Loss_Time
	dim halfsumPR_Time_Diff
	dim halfsumPR_Time_Diff_Expected
	dim halfsumPR_Calc_Point
	dim halfsumPR_Hoesu_Yul
	dim halfsumPR_Price
	
	dim half_show_day_yn
	dim half_show_night_yn
	
	dim half_show_key
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
%>
<div class="PR_Print">
<table width=960px cellpadding=0 cellspacing=0 border=0 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr style="font-weight:bold;">
	<td align=left>
		날짜 :
<%
	response.write s_PR_Work_Date
%>
		&nbsp;&nbsp;&nbsp;
		공정 :
<%
	select case strPR_Process
	case "IMD"
		response.write "IMD"
	case "SMD"
		response.write "SMD"
	case "MAN"
		response.write "수삽"
	case "ASM"
		response.write "조립"
	case "DLV"
		response.write "영업"
	end select
%>
		&nbsp;&nbsp;&nbsp;
		라인 : <%=strPR_Line%>
	</td>
</tr>
</table>
<table width=960px cellpadding=0 cellspacing=0 border=1 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr style="font-weight:bold;" bgcolor="skyblue" >
	<td width=28px>No</td>
	<td width=48px>작업</td>
	<td width=86px>파트넘버</td>
	<td width=66px>제번</td>
	<td width=52px>수량</td>
<%
	if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td width=36px>직접</td>
	<td width=36px>간접</td>
<%
	end if
%>
	<td width=46px>시작</td>
	<td width=46px>종료</td>
	<td width=40px>LOSS</td>
	<td width=50px>소요</td>
	<td width=50px>목표</td>
<%
	if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td width=50px>ST</td>
<%
	end if
%>

<%
	if instr("-IMD-SMD-","-"&strPR_Process&"-") > 0 then
%>
	<td width=40px>점수</td>
	<td width=56px>총점수</td>
<%
	end if
%>
	<td width=46px>회수율</td>
	<td width=90px>금액</td>
	<td>메모</td>
</tr>
<%
	minPR_Start_Time		= "3220"
	maxPR_End_Time			= "0820"
	
	halfminPR_Start_Time	= "3220"
	halfmaxPR_End_Time		= "0820"
	
	half_show_day_yn		= "N"
	half_show_night_yn		= "N"
	
	CNT1 = 0
	halfCNT1 = 0
	SQL = ""
	SQL	= SQL &	"select * from vwPR_List where "
	if instr(strPR_Work_Date,"between") > 0 then
		SQL = SQL & "PR_Work_Date "&strPR_Work_Date&" and "
	else
		SQL = SQL & "PR_Work_Date = '"&strPR_Work_Date&"' and "
	end if
	'SQL = SQL & "PR_Work_Date between '2008-11-01' and '2008-12-01' and "
	SQL = SQL & "PR_Process = '"&strPR_Process&"' and "
	SQL = SQL & "PR_Line = '"&strPR_Line&"' order by PR_Start_Time asc"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
%>
<tr>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td colspan=16>등록된 제조 실적이 없습니다.</td>
<%
		else
%>
	<td colspan=15>등록된 제조 실적이 없습니다.</td>
<%
		end if
%>
</tr>
<%
	else
		do until RS1.Eof
			if minPR_Start_Time > RS1("PR_Start_Time") then
				minPR_Start_Time = RS1("PR_Start_Time")
			end if
			if maxPR_End_Time < RS1("PR_End_Time") then
				maxPR_End_Time = RS1("PR_End_Time")
			end if
			
			if halfminPR_Start_Time > RS1("PR_Start_Time") then
				halfminPR_Start_Time = RS1("PR_Start_Time")
			end if
			if halfmaxPR_End_Time < RS1("PR_End_Time") then
				halfmaxPR_End_Time = RS1("PR_End_Time")
			end if
			
			sumPR_Amount					= sumPR_Amount					+ RS1("PR_Amount")
			sumPR_Worker_CNT				= sumPR_Worker_CNT				+ RS1("PR_Worker_CNT")
			sumPR_Supporter_CNT				= sumPR_Supporter_CNT			+ RS1("PR_Supporter_CNT")
			sumPR_Loss_Time					= sumPR_Loss_Time				+ RS1("PR_Loss_Time")
			sumPR_Time_Diff					= sumPR_Time_Diff				+ RS1("PR_Time_Diff")
			sumPR_Time_Diff_Expected		= sumPR_Time_Diff_Expected		+ RS1("PR_Time_Diff_Expected")
			sumPR_Calc_Point				= sumPR_Calc_Point				+ RS1("PR_Calc_Point")
			sumPR_Price						= sumPR_Price					+ RS1("PR_Price")
			sumPR_Worker_CNT_Time			= sumPR_Worker_CNT_Time			+ RS1("PR_Worker_CNT")		* RS1("PR_Time_Diff")
			sumPR_Supporter_CNT_Time		= sumPR_Supporter_CNT_Time		+ RS1("PR_Supporter_CNT")	* RS1("PR_Time_Diff")
			
			halfsumPR_Amount					= halfsumPR_Amount					+ RS1("PR_Amount")
			halfsumPR_Worker_CNT				= halfsumPR_Worker_CNT				+ RS1("PR_Worker_CNT")
			halfsumPR_Supporter_CNT				= halfsumPR_Supporter_CNT			+ RS1("PR_Supporter_CNT")
			halfsumPR_Loss_Time					= halfsumPR_Loss_Time				+ RS1("PR_Loss_Time")
			halfsumPR_Time_Diff					= halfsumPR_Time_Diff				+ RS1("PR_Time_Diff")
			halfsumPR_Time_Diff_Expected		= halfsumPR_Time_Diff_Expected		+ RS1("PR_Time_Diff_Expected")
			halfsumPR_Calc_Point				= halfsumPR_Calc_Point				+ RS1("PR_Calc_Point")
			halfsumPR_Price						= halfsumPR_Price					+ RS1("PR_Price")
			halfsumPR_Worker_CNT_Time			= halfsumPR_Worker_CNT_Time			+ RS1("PR_Worker_CNT")		* RS1("PR_Time_Diff")
			halfsumPR_Supporter_CNT_Time		= halfsumPR_Supporter_CNT_Time		+ RS1("PR_Supporter_CNT")	* RS1("PR_Time_Diff")
			
			if RS1("PR_WorkType") = "작업" then
				sumPR_Amount_ST				= sumPR_Amount_ST			+ RS1("PR_Amount")			* RS1("PR_ST")
				halfsumPR_Amount_ST			= halfsumPR_Amount_ST		+ RS1("PR_Amount")			* RS1("PR_ST")
				
				select case strPR_Line
				case "Y1"
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				case "Y2"
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				case "Y3"
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				case "RH_U"
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				case "RHSG"
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				case "RH_5"
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				case "RHAV"
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				case else
					sumPR_Amount_Point				= sumPR_Amount_Point		+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
					halfsumPR_Amount_Point			= halfsumPR_Amount_Point	+ RS1("PR_Amount")			* RS1("PR_Point") * 0.005
				end select
			end if
			
			
%>
<tr>
	<td><%=halfCNT1+1%></td>
	<td><%=RS1("PR_WorkType")%></td>
	<td><%=RS1("BOM_Sub_BS_D_No")%></td>
	<td><%=RS1("PR_Work_Order")%></td>
	
	<td align=right><%=RS1("PR_Amount")%>개&nbsp;</td>
<%
			if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td><%=RS1("PR_Worker_CNT")%>명</td>
	<td><%=RS1("PR_Supporter_CNT")%>명</td>
<%
			end if
%>
	<td><%=left(RS1("PR_Start_Time"),2)%>:<%=right(RS1("PR_Start_Time"),2)%></td>
	<td><%=left(RS1("PR_End_Time"),2)%>:<%=right(RS1("PR_End_Time"),2)%></td>
	<td align=right><%=RS1("PR_Loss_Time")%>분&nbsp;</td>
	<td align=right><%=round(RS1("PR_Time_Diff"))%>분&nbsp;</td>
<%
			if isnull(RS1("PR_Time_Diff_Expected")) then
%>
	<td align=right>-분&nbsp;</td>
<%
			else
%>
	<td align=right><%=round(RS1("PR_Time_Diff_Expected"))%>분&nbsp;</td>
<%
			end if
			if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td align=right><%=RS1("PR_ST")%>분&nbsp;</td>
<%
			end if
%>

<%
			if instr("-SMD-","-"&strPR_Process&"-") > 0 then
%>
	<td align=right><%=RS1("PR_Point")%>&nbsp;</td>
	<td align=right><%=RS1("PR_Calc_Point")%>&nbsp;</td>
<%
			elseif instr("-IMD-","-"&strPR_Process&"-") > 0 then
%>
	<td align=right>&nbsp;</td>
	<td align=right>&nbsp;</td>
<%
			end if
%>
	<td align=right><%=RS1("PR_Hoesu_Yul")%>%&nbsp;</td>
<%
			if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
				if isnumeric(RS1("PR_Price")) then
%>
	<td align=right><%=customformatcurrency(round(RS1("PR_Price")))%>&nbsp;</td>
<%
				else
%>
	<td align=center>단가미정</td>
<%
				end if
			else
%>
	<td align=right><%=customformatcurrency(RS1("PR_Calc_Point") * 4)%>&nbsp;</td>
<%			
			end if
%>
	<td align=left>&nbsp;<%=RS1("PR_Memo")%></td>
</tr>
<%
			CNT1 = CNT1 + 1
			halfCNT1 = halfCNT1 + 1
			RS1.MoveNext
			
			if RS1.Eof then
				if half_show_day_yn = "Y" and half_show_night_yn = "N" then '낮계했고, 밤계 안했다면. 밤계만 노출, 둘다 안했다면, 총계로 되므로 생략.
					half_show_key = "야간"
					half_show_night_yn = "Y"
				else
					half_show_key = "N"
				end if
			elseif RS1("PR_Start_Time") >= "2040" then '8시40분 넘어서 작업이 시작하는 경우, 첫레코드면 총계만 있으면 되므로,  불필요, 나중레코드면 낮게 표현 필요,
				if CNT1 > 1 and half_show_day_yn = "N" and minPR_Start_Time < "2040" then
					half_show_key = "주간"
					half_show_day_yn = "Y"
				else
					half_show_key = "N"
				end if
			else
				half_show_key = "N"
			end if
			
			if half_show_key <> "N" then '---------------------------------------------------------------------------------------
				
				halfsumPR_Worker_CNT				= round(halfsumPR_Worker_CNT / halfCNT1)
				halfsumPR_Supporter_CNT				= round(halfsumPR_Supporter_CNT / halfCNT1)
				halfsumPR_Time_Diff					= round(halfsumPR_Time_Diff)
				halfsumPR_Time_Diff_Expected		= round(halfsumPR_Time_Diff_Expected)
				
				halfDiff_Of_maxPR_End_Time_And_minPR_Start_Time = (left(halfmaxPR_End_Time,2)*60+right(halfmaxPR_End_Time,2))-(left(halfminPR_Start_Time,2)*60+right(halfminPR_Start_Time,2))
				halfJikJup_GongSu	= halfsumPR_Worker_CNT_Time / halfsumPR_Time_Diff * (halfDiff_Of_maxPR_End_Time_And_minPR_Start_Time)
				halfGanJup_GongSu	= halfsumPR_Supporter_CNT_Time / halfsumPR_Time_Diff * (halfDiff_Of_maxPR_End_Time_And_minPR_Start_Time)
				
				if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
					halfHoeSu_GongSu	= halfsumPR_Amount_ST
				else
					halfHoeSu_GongSu	= halfsumPR_Amount_Point
				end if
%>
<tr style="font-weight:bold;" bgcolor="#eeeeee">
	<td><%=half_show_key%></td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td align=right><%=halfsumPR_Amount%>개&nbsp;</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td><%=halfsumPR_Worker_CNT%>명</td>
	<td><%=halfsumPR_Supporter_CNT%>명</td>
<%
		end if
%>
	<td>-</td>
	<td>-</td>
	<td align=right><%=halfsumPR_Loss_Time%>분&nbsp;</td>
	<td align=right><%=halfsumPR_Time_Diff%>분&nbsp;</td>
	<td align=right><%=halfsumPR_Time_Diff_Expected%>분&nbsp;</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td align=center>-</td>
<%
		end if
%>

<%
		if instr("-SMD-","-"&strPR_Process&"-") > 0 then
%>
	<td>-</td>
	<td align=right><%=halfsumPR_Calc_Point%>&nbsp;</td>
<%
		elseif instr("-IMD-","-"&strPR_Process&"-") > 0 then
%>
	<td>-</td>
	<td align=right>&nbsp;</td>
<%
		end if
%>
	<td align=right><%=round(halfHoeSu_GongSu/(halfJikJup_GongSu+halfGanJup_GongSu)*100)%>%&nbsp;</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>	
	<td align=right><%=customformatcurrency(round(halfsumPR_Price))%>&nbsp;</td>
<%
		else
%>
	<td align=right><%=customformatcurrency(halfsumPR_Calc_Point * 4)%>&nbsp;</td>
<%
		end if
%>
	<td align=left>&nbsp;</td>
</tr>			
<%		
				halfminPR_Start_Time	= "3220"
				halfmaxPR_End_Time		= "0820"
	
				halfJikJup_GongSu				= 0
				halfGanJup_GongSu				= 0
				halfHoeSu_GongSu				= 0
	
				halfsumPR_Worker_CNT_Time		= 0
				halfsumPR_Supporter_CNT_Time	= 0
				halfsumPR_Amount_ST				= 0
				halfsumPR_Amount_Point			= 0
				
				halfsumPR_Amount				= 0
				halfsumPR_Worker_CNT			= 0
				halfsumPR_Supporter_CNT			= 0
				halfsumPR_Loss_Time				= 0
				halfsumPR_Time_Diff				= 0
				halfsumPR_Time_Diff_Expected	= 0
				halfsumPR_Calc_Point			= 0
				halfsumPR_Hoesu_Yul				= 0
				halfsumPR_Price					= 0
				
				halfCNT1 = 0
				half_show_key = "N"
			end if
			
		loop
		RS1.Close				'---------------------------------------------------------------------------------------
		
		sumPR_Worker_CNT				= round(sumPR_Worker_CNT / CNT1)
		sumPR_Supporter_CNT				= round(sumPR_Supporter_CNT / CNT1)
		sumPR_Time_Diff					= round(sumPR_Time_Diff)
		if isnull(sumPR_Time_Diff_Expected) then
			sumPR_Time_Diff_Expected		= 0
		else
			sumPR_Time_Diff_Expected		= round(sumPR_Time_Diff_Expected)
		end if
		
		Diff_Of_maxPR_End_Time_And_minPR_Start_Time = (left(maxPR_End_Time,2)*60+right(maxPR_End_Time,2))-(left(minPR_Start_Time,2)*60+right(minPR_Start_Time,2))
		JikJup_GongSu	= sumPR_Worker_CNT_Time / sumPR_Time_Diff * (Diff_Of_maxPR_End_Time_And_minPR_Start_Time)
		GanJup_GongSu	= sumPR_Supporter_CNT_Time / sumPR_Time_Diff * (Diff_Of_maxPR_End_Time_And_minPR_Start_Time)
		
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
			HoeSu_GongSu	= sumPR_Amount_ST
		else
			HoeSu_GongSu	= sumPR_Amount_Point
		end if
%>
<tr style="font-weight:bold;" bgcolor=pink>
	<td>총계</td>
	<td>-</td>
	<td>-</td>
	<td>-</td>
	<td align=right><%=sumPR_Amount%>개&nbsp;</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td><%=sumPR_Worker_CNT%>명</td>
	<td><%=sumPR_Supporter_CNT%>명</td>
<%
		end if
%>
	<td><%=left(minPR_Start_Time,2)%>:<%=right(minPR_Start_Time,2)%></td>
	<td><%=left(maxPR_End_Time,2)%>:<%=right(maxPR_End_Time,2)%></td>
	<td align=right><%=sumPR_Loss_Time%>분&nbsp;</td>
	<td align=right><%=sumPR_Time_Diff%>분&nbsp;</td>
	<td align=right><%=sumPR_Time_Diff_Expected%>분&nbsp;</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>
	<td align=center>-</td>
<%
		end if
%>
<%
		if instr("-SMD-","-"&strPR_Process&"-") > 0 then
%>
	<td>-</td>
	<td align=right><%=sumPR_Calc_Point%>&nbsp;</td>
<%
		elseif instr("-IMD-","-"&strPR_Process&"-") > 0 then
%>
	<td>-</td>
	<td align=right>&nbsp;</td>
<%
		end if
%>
	<td align=right><%=round(HoeSu_GongSu/(JikJup_GongSu+GanJup_GongSu)*100)%>%&nbsp;</td>
<%
		if instr("-MAN-ASM-","-"&strPR_Process&"-") > 0 then
%>	
	<td align=right><%=customformatcurrency(round(sumPR_Price))%>&nbsp;</td>
<%
		else
%>
	<td align=right><%=customformatcurrency(sumPR_Calc_Point * 4)%>&nbsp;</td>
<%
		end if
%>
	<td align=left>&nbsp;</td>
</tr>
<%
	end if
%>
</div>
<table>
<br><br>
<%
	set RS1 = nothing
	set RS2 = nothing
end sub
%>

<%
sub Common_PR_List_DLV(strPR_Work_Date, strPR_Process, strPR_Line)
	
	dim SQL
	dim RS1
	dim RS2
	dim CNT1
	dim halfCNT1
	
	dim Diff_Of_maxPR_End_Time_And_minPR_Start_Time
	dim halfDiff_Of_maxPR_End_Time_And_minPR_Start_Time
		
	dim minPR_Start_Time
	dim maxPR_End_Time
	
	dim halfminPR_Start_Time
	dim halfmaxPR_End_Time
	
	dim sumPR_Worker_CNT_Time
	dim sumPR_Supporter_CNT_Time
	dim sumPR_Amount_ST
	dim sumPR_Amount_Point
	
	dim sumPR_Amount
	dim sumPR_Worker_CNT
	dim sumPR_Supporter_CNT
	dim sumPR_Loss_Time
	dim sumPR_Time_Diff
	dim sumPR_Time_Diff_Expected
	dim sumPR_Calc_Point
	dim sumPR_Hoesu_Yul
	dim sumPR_Price
	
	dim halfsumPR_Worker_CNT_Time
	dim halfsumPR_Supporter_CNT_Time
	dim halfsumPR_Amount_ST
	dim halfsumPR_Amount_Point
	
	dim halfsumPR_Amount
	dim halfsumPR_Worker_CNT
	dim halfsumPR_Supporter_CNT
	dim halfsumPR_Loss_Time
	dim halfsumPR_Time_Diff
	dim halfsumPR_Time_Diff_Expected
	dim halfsumPR_Calc_Point
	dim halfsumPR_Hoesu_Yul
	dim halfsumPR_Price
	
	dim half_show_day_yn
	dim half_show_night_yn
	
	dim half_show_key
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
%>
<div class="PR_Print">
<table width=1000px cellpadding= cellspacing=0 border=0>
<tr>
	<td width=10px></td>
	<td align=left>
		<table width=600px cellpadding=0 cellspacing=0 border=0 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
		<tr style="font-weight:bold;">
			<td align=left>
				날짜 :
<%
	response.write s_PR_Work_Date
%>
					&nbsp;&nbsp;&nbsp;
					공정 : 영업
					&nbsp;&nbsp;&nbsp;
					라인 : <%=strPR_Line%>
			</td>
		</tr>
		</table>

		<table width=600px cellpadding=0 cellspacing=0 border=1 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
		<tr style="font-weight:bold;" bgcolor="skyblue" >
			<td width=28px>No</td>
			<td width=86px>파트넘버</td>
			<td width=66px>제번</td>
			<td width=52px>수량</td>
			<td width=90px>금액</td>
			<td>메모</td>
		</tr>
	<%
		CNT1 = 0
		halfCNT1 = 0
		SQL = ""
		SQL	= SQL &	"select * from vwPR_List where "
		if instr(strPR_Work_Date,"between") > 0 then
			SQL = SQL & "PR_Work_Date "&strPR_Work_Date&" and "
		else
			SQL = SQL & "PR_Work_Date = '"&strPR_Work_Date&"' and "
		end if
		SQL = SQL & "PR_Process = '"&strPR_Process&"' and "
		SQL = SQL & "PR_Line = '"&strPR_Line&"' order by PR_Start_Time asc"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
	%>
		<tr>
			<td colspan=6>등록된 납품 실적이 없습니다.</td>
		</tr>
	<%
		else
			do until RS1.Eof
				sumPR_Price						= sumPR_Price					+ RS1("PR_Price")
	%>
		<tr>
			<td><%=halfCNT1+1%></td>
			<td><%=RS1("BOM_Sub_BS_D_No")%></td>
			<td><%=RS1("PR_Work_Order")%></td>
			<td align=right><%=RS1("PR_Amount")%>개&nbsp;</td>
	<%
				if isnumeric(RS1("PR_Price")) then
	%>
			<td align=right><%=customformatcurrency(round(RS1("PR_Price")))%>&nbsp;</td>
	<%
				else
	%>
			<td align=center>단가미정</td>
	<%
				end if
	%>
			<td align=left>&nbsp;<%=RS1("PR_Memo")%></td>
		</tr>
	<%
				CNT1 = CNT1 + 1
				halfCNT1 = halfCNT1 + 1
				RS1.MoveNext
			loop
		end if
	%>
		<tr style="font-weight:bold;" bgcolor=pink>
			<td>총계</td>
			<td>-</td>
			<td>-</td>
			<td align=right><%=sumPR_Amount%>개&nbsp;</td>
			<td align=right><%=customformatcurrency(round(sumPR_Price))%>&nbsp;</td>
			<td align=left>&nbsp;</td>
		</tr>
		</div>
		</table>
	</td>
</tr>
</table>
<br><br>
<%
	set RS1 = nothing
	set RS2 = nothing
end sub
%>
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->