<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
  
<script language="javascript">
function SheetPrint()
{
	var strURL = "";
	var strP_Work_Type = "";

	if (frmSearch.s_P_Work_Type[0].checked)
		strP_Work_Type += "IMD, ";
	if (frmSearch.s_P_Work_Type[1].checked)
		strP_Work_Type += "SMD, ";
	if (frmSearch.s_P_Work_Type[2].checked)
		strP_Work_Type += "MAN, ";
	if (frmSearch.s_P_Work_Type[3].checked)
		strP_Work_Type += "ASM, ";
	if (frmSearch.s_P_Work_Type[4].checked)
		strP_Work_Type += "N/A";

	strURL += "b_parts_out_sheet_print.asp?s_BOM_Sub_BS_D_No="	+ frmSearch.s_BOM_Sub_BS_D_No.value;
	strURL += "&s_BQ_Qty="										+ frmSearch.s_BQ_Qty.value;
	strURL += "&s_P_Work_Type="									+ strP_Work_Type;
	strURL += "&part="											+ "<%=Request("part")%>";
	if(confirm("확인을 클릭하신 후 잠시기다리시면\n인쇄 대화상자가 뜹니다."))
	{
		window.open(strURL,"PartsOutSheet","height=100px,width=100px,top=2000px,lef=2000px,status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes");
	}
}
</script>

<%
dim RS1
dim RS2
dim SQL
dim CNT1

dim BOM_Sub_BS_Code
dim s_BOM_Sub_BS_D_No
dim s_BQ_Qty
dim s_P_Work_Type

dim BQ_Order
dim BQ_Use_YN
dim P_Work_Type
dim BQ_Remark
dim BQ_CHECKSUM
dim BQ_Qty
dim P_P_No
dim BQ_P_Desc
dim BQ_P_Spec
dim BQ_P_Maker
dim BS_Confirm_YN

dim strBM_WType
dim strBM_Hyungsang
dim bShow

dim BOM_Sub_BS_D_No

s_BOM_Sub_BS_D_No	= ucase(Request("s_BOM_Sub_BS_D_No"))
s_BQ_Qty			= Request("s_BQ_Qty")
s_P_Work_Type		= Request("s_P_Work_Type")

if s_BQ_Qty = "" then
	s_BQ_Qty = 1
end if

dim B_D_No
dim strMaker2

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")



call BOMSub_Guide()


'BOM_CheckSum 시작 
dim strBOM_CheckSum
dim strSW_PNO
dim dicCheckSum
set dicCheckSum =Server.CreateObject("Scripting.Dictionary")

strBOM_CheckSum = ""
SQL = "select top 1 * "
SQL = SQL & "	from tblBOM_CheckSum "
SQL = SQL & "	where "
SQL = SQL & "		BOM_Sub_BS_D_No = '"&s_BOM_Sub_BS_D_No&"' "
SQL = SQL & "	order by BC_Apply_Date desc "
RS2.Open SQL,sys_DBCon
if not(RS2.Eof or RS2.Bof) then
	strSW_PNO = cstr(trim(RS2("MC_Merge_SW_PNO")))
	if strSW_PNO <> "" then
		strBOM_CheckSum = cstr(RS2("MC_Merge_CheckSum"))
		if dicCheckSum.Exists(strSW_PNO) then
			dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
		else
			dicCheckSum.Add strSW_PNO, strBOM_CheckSum
		end if	
	end if
	strSW_PNO = cstr(trim(RS2("MICOM1_SW_PNO")))
	if strSW_PNO <> "" then
		strBOM_CheckSum = cstr(RS2("MICOM1_CheckSum"))
		if dicCheckSum.Exists(strSW_PNO) then
			dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
		else
			dicCheckSum.Add strSW_PNO, strBOM_CheckSum
		end if	
	end if
	strSW_PNO = cstr(trim(RS2("MICOM2_SW_PNO")))
	if strSW_PNO <> "" then
		strBOM_CheckSum = cstr(RS2("MICOM2_CheckSum"))
		if dicCheckSum.Exists(strSW_PNO) then
			dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
		else
			dicCheckSum.Add strSW_PNO, strBOM_CheckSum
		end if	
	end if
	strSW_PNO = cstr(trim(RS2("EEPROM1_SW_PNO")))
	if strSW_PNO <> "" then
		strBOM_CheckSum = cstr(RS2("EEPROM1_CheckSum"))
		if dicCheckSum.Exists(strSW_PNO) then
			dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
		else
			dicCheckSum.Add strSW_PNO, strBOM_CheckSum
		end if	
	end if
	strSW_PNO = cstr(trim(RS2("EEPROM2_SW_PNO")))
	if strSW_PNO <> "" then
		strBOM_CheckSum = cstr(RS2("EEPROM2_CheckSum"))
		if dicCheckSum.Exists(strSW_PNO) then
			dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
		else
			dicCheckSum.Add strSW_PNO, strBOM_CheckSum
		end if	
	end if
end if
RS2.Close
'BOM_CheckSum 끝
%>

<img src="/img/blank.gif" width=1px height=5px><br>
<%
if Request("part") = "QC" then
%>
<table width=960px cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr>
	<td width=100% align=right style="font-size:20px;">
		<table class="pi_print_2" width=200px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
		<tr bgcolor=white>
			<td width=30px rowspan=2>결<br>재</td>
			<td>담 당</td>
			<td>검 토</td>
			<td>검 토</td>
			<td>승 인</td>
		</tr>
		<tr bgcolor=white height=40px>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
elseif Request("part") = "ChulGo" then
%>
<h4>출고요청서</h4>
<%
end if
%>

<%
SQL = "select B_D_No from tbBOM where B_Code = (select top 1 BOM_B_Code from tbBOM_Sub where BS_D_No = '"&s_BOM_Sub_BS_D_No&"')"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	B_D_No = RS1("B_D_No")
end if
RS1.Close

SQL = "select BS_Code, BS_Confirm_YN from tbBom_Sub where BS_D_No = '"&s_BOM_Sub_BS_D_No&"' and BOM_B_Code  = (select max(B_Code) from tbBOM where B_Code=BOM_B_Code and B_Version_Current_YN = 'Y')"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	BOM_Sub_BS_Code = RS1("BS_Code")
	BS_Confirm_YN = RS1("BS_Confirm_YN")
end if
RS1.Close
%>

<table width=960px cellpadding=1 cellspacing=0 border=1 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<form name="frmSearch" action="b_parts_out_sheet.asp" method="post">
<input type="hidden" name="part" value="<%=Request("part")%>">
<tr>
	<td width=70px>모델</td>
	<td width=410px>
		<table width=100% cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td width=55px>&nbsp;</td>
			<td width=200px align=center><input type="text" name="s_BOM_Sub_BS_D_No" value="<%=s_BOM_Sub_BS_D_No%>" readonly onclick="javascript:show_BOMSub_Guide(this,'frmSearch',0);" style="width:90px;text-align:center;font-size:20px;width:183px">&nbsp;</td>
			<td width=50px align=left><%=Make_S_BTN("조회","javascript:frmSearch.submit();","bom_editor")%></td>
			<td width=50px align=left><%=Make_S_BTN("인쇄","javascript:SheetPrint();","bom_editor")%></td>
			<td width=55px>&nbsp;</td>
		</tr>
		</table>

	</td>
	<td width=70px>날짜</td>
	<td width=165px style="font-size:20px;"><%=date()%></td>
	<td width=70px>비고</td>
	<td style="font-size:20px;"><%if BS_Confirm_YN = "Y" then%>검증완료<%else%>검증대기<%end if%></td>
</tr>
<tr height=40>
	<td>공정</td>
	<!--<td><input type="checkbox" name="s_P_Work_Type" value="IMD"<%if instr(s_P_Work_Type,"IMD") > 0 then%> checked<%end if%>>IMD&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="SMD"<%if instr(s_P_Work_Type,"SMD") > 0 then%> checked<%end if%>>SMD&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="MAN"<%if instr(s_P_Work_Type,"MAN") > 0 then%> checked<%end if%>>MAN&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="ASM"<%if instr(s_P_Work_Type,"ASM") > 0 then%> checked<%end if%>>ASM&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="N/A"<%if instr(s_P_Work_Type,"N/A") > 0 then%> checked<%end if%>>N/A
	</td>-->
	<td><input type="checkbox" name="s_P_Work_Type" value="AXIAL"<%if instr(s_P_Work_Type,"AXIAL") > 0 then%> checked<%end if%>>AXIAL&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="RADIAL"<%if instr(s_P_Work_Type,"RADIAL") > 0 then%> checked<%end if%>>RADIAL&nbsp;&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="SMD"<%if instr(s_P_Work_Type,"SMD") > 0 then%> checked<%end if%>>SMD&nbsp;&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="MAN"<%if instr(s_P_Work_Type,"MAN") > 0 then%> checked<%end if%>>MAN&nbsp;&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="ASM"<%if instr(s_P_Work_Type,"ASM") > 0 then%> checked<%end if%>>ASM
	</td>
	<td>수량</td>
<%
if Request("Part") = "QC" then
%>
	<td colspan=3><input type="text" name="s_BQ_Qty" value="<%=s_BQ_Qty%>" size=4 style="text-align:center;font-size:20px;"></td>
<%
elseif Request("Part") = "ChulGo" then
	dim POS_Code
	'SQL = "select max(POS_Code) from tbParts_Out_Sheet"
	'RS1.Open SQL,sys_DBCon
	'if RS1.Eof or RS1.Bof then
	'	POS_Code = 1
	'else
		'POS_Code = RS1(0) + 1
	'end if

	'for CNT1 = 1 to 9-len(POS_Code)
		'POS_Code = "0" & POS_Code
	'next
%>
<script type="text/javascript" src="/header/barcode.js"></script>
	<td width=135px style="font-size:20px;"><input type="text" name="s_BQ_Qty" value="<%=s_BQ_Qty%>" size=4 style="text-align:center;font-size:20px;"></td>
	<td width=70px>일련번호</td>
	<td align=left valign=top height=30>&nbsp;&nbsp;&nbsp;&nbsp;<script language=javascript>barcode("<%=POS_Code%>", 25);</script></td>
<%
end if
%>
</tr>
</form>
</table>
<br>
<%
dim BPM_PartNo
dim BPM_Memo_Dev
dim BPM_Memo_QA
if s_BOM_Sub_BS_D_No <> "" then
	if isnumeric(left(s_BOM_Sub_BS_D_No,3)) then
		BPM_PartNo = left(s_BOM_Sub_BS_D_No,10)
	else
		BPM_PartNo = left(s_BOM_Sub_BS_D_No,9)
	end if
	SQL = "select * from tbBOM_PartsOutSheet_Memo_Dev where "
	SQL = SQL & " BPM_PartNo = '"&BPM_PartNo&"' and "
	SQL = SQL & " '"&date()&"' between BPM_StartDate and BPM_EndDate "
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		BPM_Memo_Dev = ""
	else
		BPM_Memo_Dev = trim(RS1("BPM_Memo"))
	end if
	RS1.Close
	SQL = "select * from tbBOM_PartsOutSheet_Memo_QA where "
	SQL = SQL & " BPM_PartNo = '"&BPM_PartNo&"' and "
	SQL = SQL & " '"&date()&"' between BPM_StartDate and BPM_EndDate "
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		BPM_Memo_QA = ""
	else
		BPM_Memo_QA = trim(RS1("BPM_Memo"))
	end if
	RS1.Close
	
	if BPM_Memo_Dev <> "" or BPM_Memo_QA <> "" then
%>
<table width=1020px cellpadding=1 cellspacing=0 border=1 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse;font-size:10px;font-family:arial;">
<tr bgcolor=skyblue>
	<td width=1020px colspan=2>제원시트주기</td>
</tr>
<tr bgcolor=skyblue>
	<td width=510px>품질팀</td>
	<td width=510px>개발팀</td>
</tr>
<tr bgcolor=white>
	<td width=510px style="text-align:left; vertical-align:top;"><pre><%=BPM_Memo_QA%></pre></td>
	<td width=510px style="text-align:left; vertical-align:top;"><pre><%=BPM_Memo_Dev%></pre></td>
</tr>
</table>
<br>
<%
	end if
end if
%>
<table width=1020px cellpadding=1 cellspacing=0 border=1 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse;font-size:10px;font-family:arial;">
<tr bgcolor=skyblue>
	<td width=30px>번호</td>
	<td width=30px>순번</td>
	<td width=40px>공정</td>
	<td width=90px>파트넘버</td>
	<td width=160px>작업위치</td>
	<td width=40px>수량</td>
	<!--<td width=30px>출고</td>
	<td width=30px>비고</td>-->
	<td width=150px>설명</td>
	<td width=60px>체크섬</td>
	<td width=60px>형상</td>
	<td>스펙</td>
	<td width=80px>메이커</td>
</tr>
<%
if BOM_Sub_BS_Code = "" then
%>
<tr bgcolor=white>
	<td colspan=10>검색된 BOM이 없습니다.</td>
</tr>
<%
else
	SQL = 		"select "&vbcrlf
	SQL = SQL & "	BQ_Order, "&vbcrlf
	SQL = SQL & "	BQ_Use_YN, "&vbcrlf
	SQL = SQL & "	P_Work_Type, "&vbcrlf
	SQL = SQL & "	Parts_P_P_No, "&vbcrlf
	SQL = SQL & "	BQ_Remark, "&vbcrlf
	SQL = SQL & "	BQ_CHECKSUM, "&vbcrlf
	SQL = SQL & "	BQ_Qty, "&vbcrlf
	SQL = SQL & "	BQ_P_Desc, "&vbcrlf
	'SQL = SQL & "	P_Spec = case when P_Spec_Short <> '' then P_Spec_Short else P_Spec end, "&vbcrlf
	SQL = SQL & "	BQ_P_Spec, "&vbcrlf
	SQL = SQL & "	BQ_P_Maker "&vbcrlf
	SQL = SQL & "from "&vbcrlf
	SQL = SQL & "	tbBOM_Qty "&vbcrlf
	SQL = SQL & "	left outer join "&vbcrlf
	SQL = SQL & "	tbParts "&vbcrlf
	SQL = SQL & "	on Parts_P_P_No = P_P_No "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	BQ_Qty > 0 and "&vbcrlf
	'SQL = SQL & "	(BQ_Qty > 0 or right(BQ_Order,1) = 'R') and "&vbcrlf
	
	'if s_P_Work_Type <> "" then
	'	if instr(s_P_Work_Type,"N/A") > 0 then
	'		SQL = SQL & "	(CHARINDEX(P_Work_Type,'"&s_P_Work_Type&"') > 0 or P_Work_Type is null or P_Work_Type = '') and "&vbcrlf
	'	elseif instr(s_P_Work_Type,"MAN") > 0 or instr(s_P_Work_Type,"IMD") > 0 then
	'		SQL = SQL & "	(CHARINDEX(P_Work_Type,'"&s_P_Work_Type&"') > 0 or CHARINDEX(P_Work_Type,'I/M') > 0 ) and "&vbcrlf
	'	else
	'		SQL = SQL & "	CHARINDEX(P_Work_Type,'"&s_P_Work_Type&"') > 0 and "&vbcrlf
	'	end if
	'end if
	
	SQL = SQL & "	BOM_Sub_BS_D_No='"&s_BOM_Sub_BS_D_No &"' and exists (select top 1 B_Code from tbBOM where B_Version_Current_YN='Y' and B_Code = BOM_B_Code) order by BQ_Code"&vbcrlf
	'response.write SQL
	RS1.Open SQL,sys_DBCon

	CNT1 = 1
	do until RS1.Eof
		BQ_Order		= RS1("BQ_Order")
		BQ_Use_YN		= RS1("BQ_Use_YN")
		P_Work_Type		= RS1("P_Work_Type")
		P_P_No			= RS1("Parts_P_P_No")
		BQ_Remark		= RS1("BQ_Remark")
		'BQ_CHECKSUM		= RS1("BQ_CHECKSUM")
		BQ_Qty			= RS1("Bq_Qty")
		BQ_P_Desc		= RS1("BQ_P_Desc")
		BQ_P_Spec		= RS1("BQ_P_Spec")
		BQ_P_Maker		= RS1("BQ_P_Maker")
		
		BQ_CHECKSUM = dicCheckSum.Item(cstr(P_P_No))
		
		if isnull(P_Work_Type) or P_Work_Type = "" then
			P_Work_Type = "&nbsp;"
		end if

		'BQ_Qty = BQ_Qty * s_BQ_Qty
		
		'if B_D_No <> "" then
		'	SQL = "select top 1 BM_Maker "
		'	SQL = SQL & "	from tblBOM_Mask "
		'	SQL = SQL & "	where "
		'	SQL = SQL & "		BOM_Parts_BP_PNO = '"&P_P_No&"' and "
		'	SQL = SQL & "		(BM_Filter = '_' or BM_Filter like '%"&B_D_No&"%') "
		'	SQL = SQL & "	order by BM_Filter desc "
		'	RS2.Open SQL,sys_DBCon
		'	if not(RS2.Eof or RS2.Bof) then
		'		strMaker2 = RS2("BM_Maker")
		'	end if
		'	RS2.Close
		'	
		'	'BQ_P_Maker = strMaker2
		'end if
		
		strBM_WType = ""
		strBM_Hyungsang = ""
		SQL = "select top 1 BM_WType, BM_Desc, BM_Maker, BM_Hyungsang "
		SQL = SQL & "	from tblBOM_Mask "
		SQL = SQL & "	where "
		SQL = SQL & "		BOM_Parts_BP_PNO = '"&P_P_No&"' and "
		SQL = SQL & "		(BM_Filter = '_' or BM_Filter like '%"&B_D_No&"%') "
		SQL = SQL & "	order by BM_Filter desc "
		RS2.Open SQL,sys_DBCon
		if not(RS2.Eof or RS2.Bof) then
			strBM_WType = RS2("BM_WType")
			strBM_Hyungsang = RS2("BM_Hyungsang")
		end if
		RS2.Close
		
		'SUB, MAT는 제원시트에 나오면 안됨
		bShow = false
		if instr(s_P_Work_Type,"AXIAL") > 0 then
			if instr("-PCB-AXIAL-EYELET-","-"&ucase(strBM_WType)&"-") > 0 then
				bShow = true
			end if
		end if
		if instr(s_P_Work_Type,"RADIAL") > 0 then
			if instr("-PCB-RADIAL-RADIAL_OTHER-","-"&ucase(strBM_WType)&"-") > 0 then
				bShow = true
			end if
		end if
		if instr(s_P_Work_Type,"SMD") > 0 then
			if instr("-PCB-CHIP-CHIP_OTHER-PROGRAM-CONNECTOR_S-CONNECTOR_P-","-"&ucase(strBM_WType)&"-") > 0 then
				bShow = true
			end if
			if left(strBM_WType,4) = "QFP_" then
				bShow = true
			end if 
		end if
		if instr(s_P_Work_Type,"MAN") > 0 then
			if instr("-PCB-MAN-AUTO-","-"&ucase(strBM_WType)&"-") > 0 then
				bShow = true
			end if
		end if
		if instr(s_P_Work_Type,"ASM") > 0 then
			if instr("-ASSY-ASM-","-"&ucase(strBM_WType)&"-") > 0 then
				bShow = true
			end if
		end if
		if s_P_Work_Type = "" then
			bShow = true
		end if
		'SUB, MAT는 제원시트에 나오면 안됨
		
		if bShow = true then
%>
<tr bgcolor=white style="word-wrap:break-word">
	<td><%=CNT1%></td>
	<td><%=BQ_Order%></td>
	<td><%=strBM_WType%></td>
	<td><%=P_P_No%></td>
	<td><%=BQ_Remark%></td>
	<td><%=BQ_Qty%></td>
	<!--<td>&nbsp;</td>
	<td>&nbsp;</td>-->
	<td><%=BQ_P_Desc%></td>
	<td><%=BQ_CHECKSUM%></td>
	<td><%=strBM_Hyungsang%></td>
	<td><%=BQ_P_Spec%></td>
	<td><%=BQ_P_Maker%></td>
</tr>
<%
			CNT1 = CNT1 + 1
		end if
		RS1.MoveNext
	loop
	RS1.Close
end if
%>
</table>
<%
if s_BOM_Sub_BS_D_No <> "" and gM_ID="shindk_" and Request("part") <> "ChulGo" then
%>
<br>
<table width=1024px cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed" style="border:none;font-family:arial;">
<tr>
	<td align=center><span style="font-size:17px">Option Parts List</span></td>
</tr>
</table>
<table width=1074px cellpadding=1 cellspacing=0 border=1 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse;font-size:10px;font-family:arial;">
<tr>
	<td width=35px>순번</td>
	<td width=85px>파트넘버</td>
	<td width=50px>위치</td>
	<td width=50px>설명</td>
	<td width=50px>체크섬</td>
	<td width=150px>스펙</td>
	<!--<td width=100px>메이커</td>-->
<%
dim BS_D_No_Char
dim cntBOM_Sub	

SQL = "select BS_D_No from tbBOM_Sub where BOM_B_Code = (select B_Code from tbBOM where B_D_No='"&B_D_No&"' and B_Version_Current_YN = 'Y') order by BS_D_No"
RS1.Open SQL,sys_DBCon

cntBOM_Sub = 0

do until RS1.Eof
	if isnumeric(right(trim(RS1("BS_D_No")),1)) then
		BS_D_No_Char = right(trim(RS1("BS_D_No")),2)
	else
		BS_D_No_Char = right(trim(RS1("BS_D_No")),1)
	end if	
	
	cntBOM_Sub = cntBOM_Sub + 1	
%>
	<td width=23><%=BS_D_No_Char%></td>
<%
	RS1.MoveNext
loop
RS1.Close
%>
	<td></td>
</tr>
<%
dim CNT2
dim strParts_P_P_No
dim strBQ_Qty
dim strBQ_Order
dim strBQ_Remark
dim strBQ_CHECKSUM
dim strBQ_P_Desc
dim strBQ_P_Spec
dim strBQ_P_Maker
dim arrParts_P_P_No
dim arrBQ_Qty
dim arrBQ_Order
dim arrBQ_Remark
dim arrBQ_CHECKSUM
dim arrBQ_P_Desc
dim arrBQ_P_Spec
dim arrBQ_P_Maker
dim arr2Parts_P_P_No
dim arr2BQ_Qty
dim arr2BQ_Order
dim arr2BQ_Remark
dim arr2BQ_CHECKSUM
dim arr2BQ_P_Desc
dim arr2BQ_P_Spec
dim arr2BQ_P_Maker

dim strBOM_Qty

dim Show_YN
dim cntParts

dim old_Qty

SQL = "select * from tbBOM_Qty where BOM_Sub_BS_D_No like '"&B_D_No&"%' and exists (select top 1 B_Code from tbBOM where B_Version_Current_YN='Y' and B_Code = BOM_B_Code) order by BQ_Code"
'response.write SQL
RS1.Open SQL,sys_DBCon,1

dim strCurrentError_YN
strCurrentError_YN = "N"
if RS1.Eof or RS1.Bof then
	strCurrentError_YN = "Y"
elseif not(isnumeric(RS1.RecordCount)) then
	strCurrentError_YN = "Y"
elseif not(isnumeric(cntBOM_Sub)) then
	strCurrentError_YN = "Y"
elseif cntBOM_Sub = 0 then
	strCurrentError_YN = "Y"
end if

if strCurrentError_YN = "Y" then
%>
<form name="frmRedirect" action="/bom/b_parts_out_sheet.asp?part=QC" method="post">
</form>
<Script>
alert("현재적용중인 BOM이 없습니다.\n개발팀에 문의해주세요.");
frmRedirect.submit();
</script>
<%
end if

cntParts = RS1.RecordCount / cntBOM_Sub

do until RS1.Eof
	P_P_No = RS1("Parts_P_P_No")
	
	strParts_P_P_No	= strParts_P_P_No	& RS1("Parts_P_P_No")	& "||"
	strBQ_Qty		= strBQ_Qty			& RS1("BQ_Qty")			& "||"
	strBQ_Order		= strBQ_Order		& RS1("BQ_Order")		& "||"
	strBQ_Remark	= strBQ_Remark		& RS1("BQ_Remark")		& "||"
	'strBQ_CHECKSUM	= strBQ_CHECKSUM	& RS1("BQ_CHECKSUM")		& "||"
	strBQ_P_Desc	= strBQ_P_Desc		& RS1("BQ_P_Desc")	& "||"
	strBQ_P_Spec	= strBQ_P_Spec		& RS1("BQ_P_Spec")	& "||"
	'strBQ_P_Maker	= strBQ_P_Maker		& RS1("BQ_P_Maker")	& "||"
	
	strBOM_CheckSum = ""
	SQL = "select top 1 * "
	SQL = SQL & "	from tblBOM_CheckSum "
	SQL = SQL & "	where "
	SQL = SQL & "		MC_Merge_SW_PNO = '"&P_P_No&"' or "
	SQL = SQL & "		MICOM1_SW_PNO = '"&P_P_No&"' or "
	SQL = SQL & "		MICOM2_SW_PNO = '"&P_P_No&"' or "
	SQL = SQL & "		EEPROM1_SW_PNO = '"&P_P_No&"' or "
	SQL = SQL & "		EEPROM2_SW_PNO = '"&P_P_No&"' "	
	SQL = SQL & "	order by BC_Apply_Date desc "
	RS2.Open SQL,sys_DBCon
	if not(RS2.Eof or RS2.Bof) then
		if RS2("MC_Merge_SW_PNO") = P_P_No then
			strBOM_CheckSum = RS2("MC_Merge_CheckSum")
		elseif RS2("MICOM1_SW_PNO") = P_P_No then
			strBOM_CheckSum = RS2("MICOM1_CheckSum")
		elseif RS2("MICOM2_SW_PNO") = P_P_No then
			strBOM_CheckSum = RS2("MICOM2_CheckSum")
		elseif RS2("EEPROM1_SW_PNO") = P_P_No then
			strBOM_CheckSum = RS2("EEPROM1_CheckSum")
		elseif RS2("EEPROM2_SW_PNO") = P_P_No then
			strBOM_CheckSum = RS2("EEPROM2_CheckSum")
		end if
	end if
	RS2.Close
	strBQ_CHECKSUM	= strBQ_CHECKSUM	& strBOM_CheckSum	& "||"
	
	RS1.MoveNext
loop
RS1.Close
arrParts_P_P_No	= split(strParts_P_P_No,"||")
arrBQ_Qty		= split(strBQ_Qty,"||")
arrBQ_Order		= split(strBQ_Order,"||")
arrBQ_Remark	= split(strBQ_Remark,"||")
arrBQ_CHECKSUM	= split(strBQ_CHECKSUM,"||")
arrBQ_P_Desc	= split(strBQ_P_Desc,"||")
arrBQ_P_Spec	= split(strBQ_P_Spec,"||")
'arrBQ_P_Maker	= split(strBQ_P_Maker,"||")

strParts_P_P_No = ""
strBQ_Qty		= ""
strBQ_Order		= ""
strBQ_Remark	= ""
strBQ_CHECKSUM	= ""
strBQ_P_Desc	= ""
strBQ_P_Spec	= ""
'strBQ_P_Maker	= ""

for CNT1 = 1 to cntParts
	for CNT2 = CNT1-1 to ubound(arrParts_P_P_No)-1 step cntParts
		strParts_P_P_No	= strParts_P_P_No	& arrParts_P_P_No(CNT2)	& "||"
		strBQ_Qty		= strBQ_Qty			& arrBQ_Qty(CNT2)		& "||"
		strBQ_Order		= strBQ_Order		& arrBQ_Order(CNT2)		& "||"
		strBQ_Remark	= strBQ_Remark		& arrBQ_Remark(CNT2)	& "||"
		strBQ_CHECKSUM	= strBQ_CHECKSUM	& arrBQ_CHECKSUM(CNT2)	& "||"
		strBQ_P_Desc	= strBQ_P_Desc		& arrBQ_P_Desc(CNT2)	& "||"
		strBQ_P_Spec	= strBQ_P_Spec		& arrBQ_P_Spec(CNT2)	& "||"
		'strBQ_P_Maker	= strBQ_P_Maker		& arrBQ_P_Maker(CNT2)	& "||"
	next
	strParts_P_P_No	= strParts_P_P_No	& "//"
	strBQ_Qty		= strBQ_Qty			& "//"
	strBQ_Order		= strBQ_Order		& "//"
	strBQ_Remark	= strBQ_Remark		& "//"
	strBQ_CHECKSUM	= strBQ_CHECKSUM	& "//"
	strBQ_P_Desc	= strBQ_P_Desc		& "//"
	strBQ_P_Spec	= strBQ_P_Spec		& "//"
	'strBQ_P_Maker	= strBQ_P_Maker		& "//"
next

arrParts_P_P_No	= split(strParts_P_P_No,"//")
arrBQ_Qty		= split(strBQ_Qty,"//")
arrBQ_Order		= split(strBQ_Order,"//")
arrBQ_Remark	= split(strBQ_Remark,"//")
arrBQ_CHECKSUM	= split(strBQ_CHECKSUM,"//")
arrBQ_P_Desc	= split(strBQ_P_Desc,"//")
arrBQ_P_Spec	= split(strBQ_P_Spec,"//")
'arrBQ_P_Maker	= split(strBQ_P_Maker,"//")

for CNT1 = 0 to ubound(arrParts_P_P_No)-1
	arr2Parts_P_P_No	= split(arrParts_P_P_No(CNT1),"||")
	arr2BQ_Qty			= split(arrBQ_Qty(CNT1),"||")
	arr2BQ_Order		= split(arrBQ_Order(CNT1),"||")
	arr2BQ_Remark		= split(arrBQ_Remark(CNT1),"||")
	arr2BQ_CHECKSUM		= split(arrBQ_CHECKSUM(CNT1),"||")
	arr2BQ_P_Desc		= split(arrBQ_P_Desc(CNT1),"||")
	arr2BQ_P_Spec		= split(arrBQ_P_Spec(CNT1),"||")
	'arr2BQ_P_Maker		= split(arrBQ_P_Maker(CNT1),"||")
%>


<%
	Show_YN = "N"
	strBOM_Qty = ""
	for CNT2 = 0 to ubound(arr2Parts_P_P_No)-1
		
		if CNT2 > 0 and arr2BQ_Qty(CNT2) <> old_Qty then
			Show_YN = "Y"
		end if
				
		if arr2BQ_Qty(CNT2) = 0 then
			strBOM_Qty = strBOM_Qty & "<td>&nbsp;</td>"
		else
			strBOM_Qty = strBOM_Qty & "<td>"&arr2BQ_Qty(CNT2)&"</td>"
		end if
		
		old_Qty = arr2BQ_Qty(CNT2)
	next
	
	if Show_YN = "Y" then
		'SQL = "select * from tbParts where P_P_No = '"&arr2Parts_P_P_No(0)&"'"
		'RS1.Open SQL,sys_DBCon
		
		
%>
<tr>
	<td><%=arr2BQ_Order(0)%></td>
	<td><%=arr2Parts_P_P_No(0)%></td>
	<td><%=arr2BQ_Remark(0)%></td>
	<td><%=arr2BQ_P_Desc(0)%></td>
	<td><%=arr2BQ_CHECKSUM(0)%></td>
	<td><%=arr2BQ_P_Spec(0)%></td>
	<!--<td><%rem =RS1("BQ_P_Maker")%></td>-->
<%
		'RS1.Close
		'response.write strBOM_Qty
%>
	<td>&nbsp;</td>
</tr>
<%
	end if
	
	strBOM_Qty	= ""
	Show_YN		= "N"
next
%>

</table>
<%

end if


set RS1 = nothing
set RS2 = nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->