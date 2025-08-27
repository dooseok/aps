<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim SQL
dim RS1
dim RS2

dim CNT1
dim CNT2
dim CNT3

dim s_Views_Others_YN
dim s_Edit_Process
dim s_Min_LPD_Input_Date
dim s_Diff_LPD_Input_Date
dim Max_LPD_Input_Date

dim LP_Tool_Type

s_Views_Others_YN		= Request("s_Views_Others_YN")
s_Edit_Process			= Request("s_Edit_Process")
s_Min_LPD_Input_Date	= Request("s_Min_LPD_Input_Date")
s_Diff_LPD_Input_Date	= Request("s_Diff_LPD_Input_Date")

if s_Min_LPD_Input_Date = "" then
	s_Min_LPD_Input_Date = date()
end if
if s_Diff_LPD_Input_Date = "" then
	s_Diff_LPD_Input_Date = 27
end if
if s_Views_Others_YN = "" then
	s_Views_Others_YN = "N"
end if

dim arrBOM_Sub_BS_D_No
dim BOM_Sub_BS_D_No
dim Date_Qty

dim strBG_Color
dim strDate

dim strDiff_MPD_Date
dim strMPD_Process
dim strMPD_Qty_Sum
dim arrDiff_MPD_Date
dim arrMPD_Process
dim arrMPD_Qty_Sum

dim bgMSE_Plan_Editor
dim strCall_MSE_Plan_Editor
dim strQty
dim strHidden
dim strBlank
dim strEdit_Process
dim strOther_Process
dim Edit_Process_Exist_YN
dim Other_Process_Exist_YN

dim Flag_YN

dim arrInputSelectG_1
dim arrInputSelect_1
dim arrInputSelectG_2
dim arrInputSelect_2

Max_LPD_Input_Date = dateadd("d",s_Diff_LPD_Input_Date,s_Min_LPD_Input_Date)

dim strDateQty
dim arrDateQty
for CNT1 = 0 to s_Diff_LPD_Input_Date
	strDateQty = strDateQty & ","
next
arrDateQty = split(strDateQty,",")
for CNT1 = 0 to ubound(arrDateQty)
	arrDateQty(CNT1) = 0
next

dim strTotalDateQty
dim arrTotalDateQty
for CNT1 = 0 to s_Diff_LPD_Input_Date
	strTotalDateQty = strTotalDateQty & ","
next
arrTotalDateQty = split(strTotalDateQty,",")
for CNT1 = 0 to ubound(arrTotalDateQty)
	arrTotalDateQty(CNT1) = 0
next
%>
<script language="javascript">
function frmDate_Search_Check()
{
	Show_Progress();
	frmDate_Search.submit();
}
</script>
<table border=0 cellspacing=1 cellpadding=0 width=520px bgcolor="#999999" align=center class="LGE_Plan">
<form name="frmDate_Search" action="lp_view_summary_price.asp" method="post">
<tr height=25px>
	<td bgcolor=white>
		<table border=0 cellspacing=2 cellpadding=0 width=100% bgcolor="#ffffff">
		<tr>
			<td width=100px></td>
			<td width=5px>&nbsp;</td>
			<td width=30px align=right>기간</td>
			<td width=180px align=center>
				<input type="text" name="s_Min_LPD_Input_Date" size=10 class="input" readonly value="<%=s_Min_LPD_Input_Date%>" onclick="Calendar_D(document.frmDate_Search.s_Min_LPD_Input_Date);">
				부터
				<select name="s_Diff_LPD_Input_Date">
<%
for CNT1 = 1 to 60
%>
				<option value="<%=CNT1%>"<%if int(s_Diff_LPD_Input_Date)=CNT1 then%> selected<%end if%>><%=CNT1+1%></option>
<%
next
%>
				</select>일간
			</td>
			<td width=15px></td>
			<!--
			<td width=70px align=right>공정</td>
			<td width=50px align=left>
				<select name="s_Edit_Process">
				<option value=""<%if s_Edit_Process="" then%> selected<%end if%>>-선택-</option>
				<option value="IMD"<%if s_Edit_Process="IMD" then%> selected<%end if%>>IMD</option>
				<option value="SMD"<%if s_Edit_Process="SMD" then%> selected<%end if%>>SMD</option>
				<option value="MAN"<%if s_Edit_Process="MAN" then%> selected<%end if%>>MAN</option>
				<option value="ASM"<%if s_Edit_Process="ASM" then%> selected<%end if%>>ASM</option>
				</select>
			<td width=15px></td>
			<td width=70px align=right>타공정보기</td>
			<td width=15px align=left>
				<input type="checkbox" name="s_Views_Others_YN" value="Y"<%if s_Views_Others_YN = "Y" then%> checked<%end if%>>
			</td>
			-->
			<td width=50px><%=Make_S_BTN("조회","javascript:frmDate_Search_Check();","")%></td>
			<td width=5px>&nbsp;</td>
			<td width=100px></td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<br>
<table width="<%=420+(30*(s_Diff_LPD_Input_Date+1))%>"px cellpadding=0 cellspacing=0 border=0 bgcolor="#999999" class="LGE_Plan">
<tr bgcolor=dimgray>
	<td width=100px style="color:white"><b>PART NO</b></td>
	<td width=100px style="color:white"><b>TOOL</b></td>
<%	
for CNT1 = 0 to s_Diff_LPD_Input_Date
	strDate = dateadd("d",CNT1,s_Min_LPD_Input_Date)
	strDate = right(strDate,2)
%>
	<td width=40px style="color:white"><b><%=strDate%></b></td>
<%
next
%>
	<td width=40px align=right style="color:white"><b>QTY</b>&nbsp;</td>
	<td width=80px align=right style="color:white"><b>PRICE</b>&nbsp;&nbsp;&nbsp;</td>
	<td width=100px align=right style="color:white"><b>SUM</b>&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<%
SQL = 		"select "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
for CNT1 = 0 to s_Diff_LPD_Input_Date
	SQL = SQL & "	Date_Qty_"&CNT1&" = ISNULL(sum(case when LPD_Input_Date = '"&dateadd("d",CNT1,s_Min_LPD_Input_Date)&"' then LPD_Input_Qty end),0), "&vbcrlf
next
SQL = SQL & "	LP_Tool_Type = TI_Type "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	tbLGE_Plan_Date "&vbcrlf
SQL = SQL & "	left outer join "&vbcrlf
SQL = SQL & "	vwLM_List "&vbcrlf
SQL = SQL & "	on LGE_Plan_LP_Model = LM_Name "&vbcrlf
SQL = SQL & "	left outer join "&vbcrlf
SQL = SQL & "	tbTool_Info "&vbcrlf
SQL = SQL & "	on TI_Name = (select top 1 LP_Tool from tbLGE_Plan where LP_Model = LM_Name) "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No is not null	and "&vbcrlf
SQL = SQL & "	LM_Company = 'MSE' and "&vbcrlf
SQL = SQL & "	LPD_Input_Date between '"&s_Min_LPD_Input_Date&"' and '"&Max_LPD_Input_Date&"' "&vbcrlf
SQL = SQL & "group by "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
SQL = SQL & "	TI_Type "&vbcrlf
SQL = SQL & "order by LP_Tool_Type, "&vbcrlf
for CNT1 = 0 to s_Diff_LPD_Input_Date
	SQL = SQL & "	Date_Qty_"&CNT1&" desc, "&vbcrlf
next
SQL = SQL & "	BOM_Sub_BS_D_No asc "&vbcrlf

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL,sys_DBCon

dim old_LP_Tool_Type

dim QTY_BOM_Sub_BS_D_No
dim Price_BOM_Sub_BS_D_No
dim CAL_BOM_Sub_BS_D_No

dim Sum_QTY_BOM_Sub_BS_D_No
dim Sum_CAL_BOM_Sub_BS_D_No

dim Total_QTY_BOM_Sub_BS_D_No
dim Total_CAL_BOM_Sub_BS_D_No

old_LP_Tool_Type = RS1("LP_Tool_Type")
do until RS1.Eof
	BOM_Sub_BS_D_No = RS1("BOM_Sub_BS_D_No")
	BOM_Sub_BS_D_No = replace(BOM_Sub_BS_D_No,chr(13),"<br>")
	BOM_Sub_BS_D_No = replace(BOM_Sub_BS_D_No,"<br><br><br>","")
	BOM_sub_BS_D_No = replace(BOM_Sub_BS_D_No,"<br><br>","")
	
	LP_Tool_Type	= RS1("LP_Tool_Type")
	
	if right(BOM_Sub_BS_D_No,4) = "<br>" then
		BOM_Sub_BS_D_No = left(BOM_Sub_BS_D_No,len(BOM_Sub_BS_D_No)-4)
	end if
	
	arrBOM_Sub_BS_D_No = split(BOM_Sub_BS_D_No,"<br>")
	
	Flag_YN = "N"
	for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)
	
		if old_LP_Tool_Type <> LP_Tool_Type then
%>
<tr bgcolor=white height=20px>
	<td></td>
	<td></td>
<%
			for CNT2 = 0 to s_Diff_LPD_Input_Date
				arrTotalDateQty(CNT2) = arrTotalDateQty(CNT2) + arrDateQty(CNT2)
				if arrDateQty(CNT2) = 0 then
					arrDateQty(CNT2) = ""
				end if
%>
	<td style="color:darkblue"><%=arrDateQty(CNT2)%></td>
<%
				arrDateQty(CNT2) = 0
			next
%>
	<td width=40px align=right style="color:darkblue"><%=Sum_QTY_BOM_Sub_BS_D_No%>&nbsp;</td>
	<td width=80px>&nbsp;</td>
	<td width=100px align=right style="color:darkblue"><%=customformatcurrency(Sum_CAL_BOM_Sub_BS_D_No)%>&nbsp;</td>
</tr>
<tr height=15px><td colspan=100 bgcolor="#ffffff"><img src="/img/blank.gif" width=1px height=15px></td></tr>
<tr bgcolor=dimgray>
	<td width=100px style="color:white"><b>PART NO</b></td>
	<td width=100px style="color:white"><b>TOOL</b></td>
<%	
		for CNT2 = 0 to s_Diff_LPD_Input_Date
			strDate = dateadd("d",CNT2,s_Min_LPD_Input_Date)
			strDate = right(strDate,2)
%>
	<td width=40px style="color:white"><b><%=strDate%></b></td>
<%
		next
%>
	<td width=40px align=right style="color:white"><b>QTY</b>&nbsp;</td>
	<td width=80px align=right style="color:white"><b>PRICE</b>&nbsp;&nbsp;&nbsp;</td>
	<td width=100px align=right style="color:white"><b>SUM</b>&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<%
			Total_QTY_BOM_Sub_BS_D_No = Total_QTY_BOM_Sub_BS_D_No + Sum_QTY_BOM_Sub_BS_D_No
			Total_CAL_BOM_Sub_BS_D_No = Total_CAL_BOM_Sub_BS_D_No + Sum_CAL_BOM_Sub_BS_D_No
			Sum_QTY_BOM_Sub_BS_D_No = 0
			Sum_CAL_BOM_Sub_BS_D_No	= 0
		end if
	
		'strDiff_MPD_Date		= ""
		'strMPD_Process			= ""
		'strMPD_Qty_Sum			= ""
		'SQL =		"select distinct"&vbcrlf
		'SQL = SQL & "	Code		= (left(convert(char,MPD_Date,121),10)+BOM_Sub_BS_D_No), "&vbcrlf
		'SQL = SQL & "	Diff_MPD_Date = datediff(day,'"&s_Min_LPD_Input_Date&"',MPD_Date), "&vbcrlf
		'SQL = SQL & "	MPD_Process, "&vbcrlf
		'SQL = SQL & "	MPD_Qty_Sum	= sum(MPD_Qty) "&vbcrlf
		'SQL = SQL & "from "&vbcrlf
		'SQL = SQL & "	tbMSE_Plan_Date "&vbcrlf
		'SQL = SQL & "where "&vbcrlf
		
		'if s_Views_Others_YN <> "Y" then
			'SQL = SQL & "	MPD_Process = '"&s_Edit_Process&"' and "&vbcrlf
		'end if
		
		'SQL = SQL & "	BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"' "&vbcrlf
		'SQL = SQL & "	group by BOM_Sub_BS_D_No, MPD_Date, MPD_Process "&vbcrlf
		'RS2.Open SQL,sys_DBCon
		'do until RS2.Eof
			'strDiff_MPD_Date	= strDiff_MPD_Date	& RS2("Diff_MPD_Date")	& "//"
			'strMPD_Process		= strMPD_Process	& RS2("MPD_Process")	& "//"
			'strMPD_Qty_Sum		= strMPD_Qty_Sum	& RS2("MPD_Qty_Sum")	& "//"
			'RS2.MoveNext
		'loop
		'RS2.Close
		'arrDiff_MPD_Date	= split(strDiff_MPD_Date,"//")
		'arrMPD_Process		= split(strMPD_Process,"//")
		'arrMPD_Qty_Sum		= split(strMPD_Qty_Sum,"//")
%>
<tr bgcolor=white height=20px>
	<td width=100px><%=arrBOM_Sub_BS_D_No(CNT1)%></td>
	<td width=100px><%=LP_Tool_Type%></td>
<%	
		for CNT2 = 0 to s_Diff_LPD_Input_Date
			'if s_Edit_Process <> "" then			'수정 모드라면 수정을 위한 공용코드를 만들어 둔다.
				'strCall_MSE_Plan_Editor	= " style='cursor:hand' onclick=""show_MSE_Plan_Editor('"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"','"&arrBOM_Sub_BS_D_No(CNT1)&"','"&dateadd("d",CNT2,s_Min_LPD_Input_Date)&"')"" "
			'end if
			'strHidden	= "<span"&strCall_MSE_Plan_Editor&" id='"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"' class='"&s_Edit_Process&"' style='display:none;'>&nbsp;</span>"
			'strBlank	= "<span"&strCall_MSE_Plan_Editor&" id='"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"' class='BLANK'>&nbsp;</span>"
		
			Date_Qty = RS1("Date_Qty_"&CNT2)
			if Date_Qty = 0 then
				Date_Qty = ""
			end if
			strQty = Date_Qty
			
			'strQty = "<span"&strCall_MSE_Plan_Editor&" class='LGE_Due'>"&Date_Qty&"</span>"
			
			'Edit_Process_Exist_YN	= "N"
			'Other_Process_Exist_YN	= "N"	
			'strEdit_Process			= ""						'한 셀용 변수 초기화
			'strOther_Process		= ""
			
			'for CNT3 = 0 to ubound(arrDiff_MPD_Date) - 1	'조회된 데이터를 검색
				'if arrDiff_MPD_Date(CNT3) = cstr(CNT2) then			'그 중 현재 셀에 표시할 것이 잇다면.
					'if arrMPD_Process(CNT3) = s_Edit_Process then	'수정 대상 공정이면,
						'Edit_Process_Exist_YN	= "Y"
						'strEdit_Process = strEdit_Process & "<span"&strCall_MSE_Plan_Editor&" id='"&arrBOM_Sub_BS_D_No(CNT1)&"_"&CNT2&"' class='"&arrMPD_Process(CNT3)&"'>"&arrMPD_Qty_Sum(CNT3)&"</span>"
					'else											'수정 대상 공정이 아니면,
						'Other_Process_Exist_YN	= "Y"				
						'strOther_Process = strOther_Process & "<span"&strCall_MSE_Plan_Editor&" class='"&arrMPD_Process(CNT3)&"'>"&arrMPD_Qty_Sum(CNT3)&"</span>"
					'end if
				'end if
			'next	
			
			'if Date_Qty = "" then				'납기정보 없음
				'if Edit_Process_Exist_YN = "Y" then				'수정대상 공정 있음
					'strQty = strEdit_Process & strOther_Process
				'elseif Other_Process_Exist_YN = "Y" then		'다른 공정만 있음
					'strQty = strHidden & strEdit_Process & strOther_Process 
				'else											'공정정보 없음
					'strQty = strBlank
				'end if
			'elseif strQty <> "" then						'납기정보 있음
				'if Edit_Process_Exist_YN = "Y" then				'수정대상 공정 있음
					'strQty = strEdit_Process & strOther_Process & strQty
				'elseif Other_Process_Exist_YN = "Y" then		'다른 공정만 있음
					'strQty = strHidden & strEdit_Process & strOther_Process & strQty
				'else											'공정정보 없음
					'strQty = strHidden & strQty
				'end if
			'end if	
			
			if weekday(dateadd("d",CNT2,s_Min_LPD_Input_Date)) = 1 then
				strBG_Color = "pink"
			elseif weekday(dateadd("d",CNT2,s_Min_LPD_Input_Date)) = 7 then
				strBG_Color = "skyblue"
			elseif CNT2 mod 2 = 1 then
				strBG_Color = "#ffffff"
			else
				strBG_Color = "#e3e3e3"
			end if
			
			if isnumeric(strQty) then
				QTY_BOM_Sub_BS_D_No = QTY_BOM_Sub_BS_D_No + strQty
				
				if Flag_YN = "N" then
					arrDateQty(CNT2) = arrDateQty(CNT2) + strQty
				end if
			end if
%>
	<td width=40px bgcolor="<%=strBG_Color%>"><%=strQty%></td>
<%
		next
	
		SQL = "select top 1 BP_Price from tbBOM_Price where BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"' and BP_Currency = 'KRW' and BP_Apply_Date <= '"&s_Min_LPD_Input_Date&"' order by BP_Reg_Date desc"
		RS2.Open SQL,sys_DBCon
		if RS2.Eof or RS2.Bof then
			Price_BOM_Sub_BS_D_No = 0
		else
			Price_BOM_Sub_BS_D_No = round(RS2("BP_Price"),0)
		end if
		RS2.Close
		
		CAL_BOM_Sub_BS_D_No = Price_BOM_Sub_BS_D_No*QTY_BOM_Sub_BS_D_No
		
		if Flag_YN = "N" then
			Flag_YN = "Y"
			Sum_QTY_BOM_Sub_BS_D_No		= Sum_QTY_BOM_Sub_BS_D_No	+ QTY_BOM_Sub_BS_D_No
			Sum_CAL_BOM_Sub_BS_D_No		= Sum_CAL_BOM_Sub_BS_D_No	+ CAL_BOM_Sub_BS_D_No
		end if
%>
	<td width=40px bgcolor="<%=strBG_Color%>" align=right><%=QTY_BOM_Sub_BS_D_No%>&nbsp;</td>
	<td width=80px bgcolor="<%=strBG_Color%>" align=right><%=customformatcurrency(Price_BOM_Sub_BS_D_No)%>&nbsp;&nbsp;</td>
	<td width=100px bgcolor="<%=strBG_Color%>" align=right><%=customformatcurrency(CAL_BOM_Sub_BS_D_No)%>&nbsp;</td>
</tr>
<%
		QTY_BOM_Sub_BS_D_No = 0
		
		if CNT1 < ubound(arrBOM_Sub_BS_D_No) then
%>
<tr height=1px><td colspan=100 bgcolor="pink"><img src="/img/blank.gif" width=1px height=1px></td></tr>
<%			
		end if
	
		old_LP_Tool_Type = LP_Tool_Type
	next
%>
<tr height=1px><td colspan=100 bgcolor="#333333><img src="/img/blank.gif" width=1px height=1px></td></tr>
<%
	RS1.MoveNext
loop
RS1.Close
%>

<tr bgcolor=white height=20px>
	<td></td>
	<td></td>
<%
			for CNT2 = 0 to s_Diff_LPD_Input_Date
				arrTotalDateQty(CNT2) = arrTotalDateQty(CNT2) + arrDateQty(CNT2)
				if arrDateQty(CNT2) = 0 then
					arrDateQty(CNT2) = ""
				end if
%>
	<td style="color:darkblue"><%=arrDateQty(CNT2)%></td>
<%
				arrDateQty(CNT2) = 0
			next
			
			Total_QTY_BOM_Sub_BS_D_No = Total_QTY_BOM_Sub_BS_D_No + Sum_QTY_BOM_Sub_BS_D_No
			Total_CAL_BOM_Sub_BS_D_No = Total_CAL_BOM_Sub_BS_D_No + Sum_CAL_BOM_Sub_BS_D_No
%>
	<td width=40px align=right style="color:darkblue"><%=Sum_QTY_BOM_Sub_BS_D_No%>&nbsp;</td>
	<td width=80px>&nbsp;</td>
	<td width=100px align=right style="color:darkblue"><%=customformatcurrency(Sum_CAL_BOM_Sub_BS_D_No)%>&nbsp;</td>
</tr>
<tr height=15px><td colspan=100 bgcolor="#ffffff"><img src="/img/blank.gif" width=1px height=15px></td></tr>
<%
set RS2 = nothing
set RS1 = nothing
%>
<tr bgcolor=white height=20px>
	<td></td>
	<td></td>
<%
			for CNT2 = 0 to s_Diff_LPD_Input_Date
				if arrTotalDateQty(CNT2) = 0 then
					arrTotalDateQty(CNT2) = ""
				end if
%>
	<td style="color:darkblue"><%=arrTotalDateQty(CNT2)%></td>
<%
			next
%>
	<td width=40px align=right style="color:darkblue"><%=Total_QTY_BOM_Sub_BS_D_No%>&nbsp;</td>
	<td width=80px>&nbsp;</td>
	<td width=100px align=right style="color:darkblue"><%=customformatcurrency(Total_CAL_BOM_Sub_BS_D_No)%>&nbsp;</td>
</tr>
<tr height=15px><td colspan=100 bgcolor="#ffffff"><img src="/img/blank.gif" width=1px height=15px></td></tr>
</table>

<%
if s_Edit_Process <> "" then

	select case s_Edit_Process
		case "IMD"
			bgMSE_Plan_Editor = "#F2D4D4"
		case "SMD"
			bgMSE_Plan_Editor = "#D2F6C9"
		case "MAN"
			bgMSE_Plan_Editor = "#C6EBFE"
		case "ASM"
			bgMSE_Plan_Editor = "#EADAF7"
	end select
%>
<div id="divMSE_Plan_Editor" style="width:50px;height:50px;position:absolute;display:none;border:1px solid #999999;filter:alpha(opacity=90);">
<table width=100% cellpadding=0 cellspacing=0 border=0 bgcolor=white class="MSE_Plan_Editor">
<form name="frmMSE_Plan_Editor" action="inc_MSE_Plan_reg_action.asp" method="post" target="ifrmMSE_Plan_Action">
<input type="hidden" name="idDIV">
<input type="hidden" name="BOM_Sub_BS_D_No">
<input type="hidden" name="MPD_Process" value="<%=s_Edit_Process%>">
<input type="hidden" name="MPD_Date">
<input type="hidden" name="MPD_Qty_Total">
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2 align=left valign=top><img src="/img/ico_MSE_Plan_<%=s_Edit_Process%>.gif"></td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td width=40px>공정 :</td>
	<td align=left><%=s_Edit_Process%></td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td width=40px>날짜 :</td>
	<td align=left id="idMSE_Plan_Date"></td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td width=40px>총계 :</td>
	<td align=left id="idMPD_Qty_Total">
		&nbsp;
	</td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2>
		<table width=100% cellpadding=1 cellspacing=0 border=0>
		<tr>
			<td width=30px></td>
<%
	select case s_Edit_Process
	case "IMD"
		arrInputSelectG_2 = split(replace(BasicDataIMDLine,"slt>",""),";")
	case "SMD"
		arrInputSelectG_2 = split(replace(BasicDataSMDLine,"slt>",""),";")
	case "MAN"
		arrInputSelectG_2 = split(replace(BasicDataMANLine,"slt>",""),";")
	case "ASM"
		arrInputSelectG_2 = split(replace(BasicDataASMLine,"slt>",""),";")
	end select
	
	for CNT2 = 0 to ubound(arrInputSelectG_2)
		arrInputSelect_2 = split(arrInputSelectG_2(CNT2),":")
%>
			<td><%=arrInputSelect_2(0)%></td>
<%
	next
%>			
			<td>&nbsp;</td>
		</tr>
<%
	if s_Edit_Process="IMD" or s_Edit_Process="SMD" then
		arrInputSelectG_1	= split(replace(BasicDataFullTime,"slt>",""),";")
	else
		arrInputSelectG_1	= split(replace(BasicDataHalfTime,"slt>",""),";")
	end if

	for CNT1 = 0 to ubound(arrInputSelectG_1)
		arrInputSelect_1 = split(arrInputSelectG_1(CNT1),":")
%>
		<tr>
			<td>&nbsp;<%=arrInputSelect_1(0)%>&nbsp;</td>
<%
		select case s_Edit_Process
		case "IMD"
			arrInputSelectG_2 = split(replace(BasicDataIMDLine,"slt>",""),";")
		case "SMD"
			arrInputSelectG_2 = split(replace(BasicDataSMDLine,"slt>",""),";")
		case "MAN"
			arrInputSelectG_2 = split(replace(BasicDataMANLine,"slt>",""),";")
		case "ASM"
			arrInputSelectG_2 = split(replace(BasicDataASMLine,"slt>",""),";")
		end select
		
		for CNT2 = 0 to ubound(arrInputSelectG_2)
			arrInputSelect_2 = split(arrInputSelectG_2(CNT2),":")
%>
			<td><input type="text" name="<%=arrInputSelect_2(0)%>_<%=arrInputSelect_1(0)%>" value="" style="width:35px;text-align:center" maxlength=4 onkeyup="javascript:cal_MPD_Qty_Total()"></td>
<%
		next
%>
			<td>&nbsp;</td>
		</tr>
<%
	next
%>
		</table>
	</td>
</tr>
<tr bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2>
		<input type="button" value="확인" onclick="javascript:frmMSE_Plan_Editor_Check();">
		<input type="button" value="취소" onclick="javascript:hide_MSE_Plan_Editor();">
	</td>
</tr>
<tr height=3px bgcolor="<%=bgMSE_Plan_Editor%>">
	<td colspan=2><img src="/img/blank.gif" width=1px height=3px></td>
</tr>
</form>
</table>
<iframe name="ifrmMSE_Plan_Action" src="about:blank" width=0px height=0px frameborder=0></iframe>
</div>


<script language="javascript">
function cal_MPD_Qty_Total()
{
	var nTotal = 0;
<%
	if s_Edit_Process="IMD" or s_Edit_Process="SMD" then
		arrInputSelectG_1	= split(replace(BasicDataFullTime,"slt>",""),";")
	else
		arrInputSelectG_1	= split(replace(BasicDataHalfTime,"slt>",""),";")
	end if
	
		
	for CNT1 = 0 to ubound(arrInputSelectG_1)
		arrInputSelect_1 = split(arrInputSelectG_1(CNT1),":")
			
		select case s_Edit_Process
		case "IMD"
			arrInputSelectG_2 = split(replace(BasicDataIMDLine,"slt>",""),";")
		case "SMD"
			arrInputSelectG_2 = split(replace(BasicDataSMDLine,"slt>",""),";")
		case "MAN"
			arrInputSelectG_2 = split(replace(BasicDataMANLine,"slt>",""),";")
		case "ASM"
			arrInputSelectG_2 = split(replace(BasicDataASMLine,"slt>",""),";")
		end select
		
		for CNT2 = 0 to ubound(arrInputSelectG_2)
			arrInputSelect_2 = split(arrInputSelectG_2(CNT2),":")
%>
	if (frmMSE_Plan_Editor.<%=arrInputSelect_2(0)%>_<%=arrInputSelect_1(0)%>.value)
		nTotal += parseInt(frmMSE_Plan_Editor.<%=arrInputSelect_2(0)%>_<%=arrInputSelect_1(0)%>.value);
<%
		next
	next
%>
	frmMSE_Plan_Editor.MPD_Qty_Total.value = nTotal;
	
	var objDIV = document.getElementById("idMPD_Qty_Total");
	objDIV.innerHTML = nTotal;
}

function show_MSE_Plan_Editor(idDIV,strBOM_Sub_BS_D_No,strMSE_Plan_Date)
{
	frmMSE_Plan_Editor.reset();
	ifrmMSE_Plan_Action.location.href	= "inc_mse_plan_load_action.asp?MPD_Process=<%=s_Edit_Process%>&BOM_Sub_BS_D_No="+strBOM_Sub_BS_D_No+"&MPD_Date="+strMSE_Plan_Date;
	
	divMSE_Plan_Editor.style.posLeft	= event.x + 0 + document.body.scrollLeft;
	divMSE_Plan_Editor.style.posTop		= event.y + 0 + document.body.scrollTop;
	divMSE_Plan_Editor.style.display	= "block";

	frmMSE_Plan_Editor.BOM_Sub_BS_D_No.value				= strBOM_Sub_BS_D_No;
	frmMSE_Plan_Editor.MPD_Date.value						= strMSE_Plan_Date;
	
	document.getElementById("idMSE_Plan_Date").innerHTML	= strMSE_Plan_Date;
	frmMSE_Plan_Editor.idDIV.value							= idDIV;
}

function hide_MSE_Plan_Editor()
{
	frmMSE_Plan_Editor.reset();
	divMSE_Plan_Editor.style.display="none";
}

function frmMSE_Plan_Editor_Check()
{
	var objDIV = document.getElementById(frmMSE_Plan_Editor.idDIV.value);
	if (parseInt(frmMSE_Plan_Editor.MPD_Qty_Total.value) > 0)
	{
		objDIV.innerHTML		= frmMSE_Plan_Editor.MPD_Qty_Total.value;
		objDIV.style.display	= "block";
	
<%
	select case s_Edit_Process
	case "IMD"
%>
		objDIV.style.backgroundColor = "red";
<%
	case "SMD"
%>
		objDIV.style.backgroundColor = "green";
<%
	case "MAN"
%>
		objDIV.style.backgroundColor = "blue";
<%
	case "ASM"
%>
		objDIV.style.backgroundColor = "#7306C6";
<%
	end select
%>
	}
	else
	{
		objDIV.style.display			= "none";
		objDIV.innerHTML				= "";
		objDIV.style.backgroundColor	= "transparent";
	}
	frmMSE_Plan_Editor.submit();
}

function Load_frmMSE_Plan_Editor(strMPD_Line,strMPD_Time,strMPD_Qty)
{
	objForm = eval("frmMSE_Plan_Editor."+strMPD_Line+"_"+strMPD_Time)
	objForm.value = strMPD_Qty;
}
</script>
<%
end if
%>


<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->