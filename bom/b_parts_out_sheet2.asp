<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
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
dim BQ_Qty
dim P_P_No
dim BQ_P_Desc
dim BQ_P_Spec
dim BQ_P_Maker

dim BOM_B_D_No
dim BOM_Sub_BS_D_No

s_BOM_Sub_BS_D_No	= ucase(Request("s_BOM_Sub_BS_D_No"))
s_BQ_Qty			= Request("s_BQ_Qty")
s_P_Work_Type		= Request("s_P_Work_Type")

if s_BQ_Qty = "" then
	s_BQ_Qty = 1
end if

call BOMSub_Guide()
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
			<td width=30px rowspan=2>결<br>제</td>
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
end if
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
	<td width=410px style="font-size:20px;"><%=date()%></td>
</tr>
<tr>
	<td>공정</td>
	<td><input type="checkbox" name="s_P_Work_Type" value="IMD"<%if instr(s_P_Work_Type,"IMD") > 0 then%> checked<%end if%>>IMD&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="SMD"<%if instr(s_P_Work_Type,"SMD") > 0 then%> checked<%end if%>>SMD&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="MAN"<%if instr(s_P_Work_Type,"MAN") > 0 then%> checked<%end if%>>MAN&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="ASM"<%if instr(s_P_Work_Type,"ASM") > 0 then%> checked<%end if%>>ASM&nbsp;
		<input type="checkbox" name="s_P_Work_Type" value="N/A"<%if instr(s_P_Work_Type,"N/A") > 0 then%> checked<%end if%>>N/A
	</td>
	<td>수량</td>
	<td><input type="text" name="s_BQ_Qty" value="<%=s_BQ_Qty%>" size=4 style="text-align:center;font-size:20px;"></td>
</tr>
</form>
</table>
<br>
<%
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select BS_Code from tbBom_Sub where BS_D_No = '"&s_BOM_Sub_BS_D_No&"' and BOM_B_Code  = (select max(B_Code) from tbBOM where B_Code=BOM_B_Code and B_Version_Current_YN = 'Y')"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	BOM_Sub_BS_Code = RS1("BS_Code")
end if
RS1.Close
%>
<table width=960px cellpadding=1 cellspacing=0 border=1 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse;font-size:10px;font-family:arial;">
<tr bgcolor=skyblue>
	<td width=30px>번호</td>
	<td width=30px>순번</td>
	<td width=40px>공정</td>
	<td width=90px>파트넘버</td>
	<td width=60px>작업위치</td>
	<td width=40px>수량</td>
	<td width=30px>출고</td>
	<td width=30px>비고</td>
	<td width=150px>설명</td>
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
	SQL = SQL & "	P_P_No, "&vbcrlf
	SQL = SQL & "	BQ_Remark, "&vbcrlf
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
	SQL = SQL & "	(BQ_Qty > 0 or right(BQ_Order,1) = 'R') and "&vbcrlf
	if s_P_Work_Type <> "" then
		if instr(s_P_Work_Type,"N/A") > 0 then
			SQL = SQL & "	(CHARINDEX(P_Work_Type,'"&s_P_Work_Type&"') > 0 or P_Work_Type is null or P_Work_Type = '') and "&vbcrlf
		elseif instr(s_P_Work_Type,"MAN") > 0 or instr(s_P_Work_Type,"IMD") > 0 then
			SQL = SQL & "	(CHARINDEX(P_Work_Type,'"&s_P_Work_Type&"') > 0 or CHARINDEX(P_Work_Type,'I/M') > 0 ) and "&vbcrlf
		else
			SQL = SQL & "	CHARINDEX(P_Work_Type,'"&s_P_Work_Type&"') > 0 and "&vbcrlf
		end if
	end if
	SQL = SQL & "	BOM_Sub_BS_Code='"&BOM_Sub_BS_Code &"' and BOM_B_Code in (select B_Code from tbBOM where B_Version_Current_YN='Y') order by BQ_Code"&vbcrlf

	RS1.Open SQL,sys_DBCon

	CNT1 = 1
	do until RS1.Eof
		BQ_Order		= RS1("BQ_Order")
		BQ_Use_YN		= RS1("BQ_Use_YN")
		P_Work_Type	= RS1("P_Work_Type")
		P_P_No			= RS1("P_P_No")
		BQ_Remark		= RS1("BQ_Remark")
		BQ_Qty			= RS1("Bq_Qty")
		BQ_P_Desc		= RS1("BQ_P_Desc")
		BQ_P_Spec		= RS1("BQ_P_Spec")
		BQ_P_Maker		= RS1("BQ_P_Maker")

		if isnull(P_Work_Type) or P_Work_Type = "" then
			P_Work_Type = "&nbsp;"
		end if

		BQ_Qty = BQ_Qty * s_BQ_Qty
%>
<tr bgcolor=white style="word-wrap:break-word">
	<td><%=CNT1%></td>
	<td><%=BQ_Order%></td>
	<td><%=P_Work_Type%></td>
	<td><%=P_P_No%></td>
	<td><%=BQ_Remark%></td>
	<td><%=BQ_Qty%></td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td><%=BQ_P_Desc%></td>
	<td><%=BQ_P_Spec%></td>
	<td><%=BQ_P_Maker%></td>
</tr>
<%
		CNT1 = CNT1 + 1
		RS1.MoveNext
	loop
	RS1.Close
end if
%>
</table>
<%
response.write s_BOM_Sub_BS_D_No
if s_BOM_Sub_BS_D_No <> "" then
	response.write s_BOM_Sub_BS_D_No
%>
<table width=1024px cellpadding=1 cellspacing=0 border=1 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse;font-size:10px;font-family:arial;">
<tr>
	<td></td>
<%
dim BS_D_No_Char
SQL = "select BS_D_No from tbBOM_Sub where BOM_B_D_No = '"&s_BOM_Sub_BS_D_No&"%' and BOM_B_Code in (select B_Code from tbBOM where B_Version_Current_YN='Y') order BS_D_No"
response.write SQL
RS1.Open SQL,sys_DBCon
for CNT = 1 to 24
	if RS1.Eof or RS1.Bof then
%>
	<td width=20>&nbsp;</td>
<%		
	else
		if isnumeric(right(trim(RS1("BS_D_No")),1)) then
			BS_D_No_Char = right(trim(RS1("BS_D_No")),2)
		else
			BS_D_No_Char = right(trim(RS1("BS_D_No")),1)
		end if		
%>
	<td width=20><%=BS_D_No_Char%></td>
<%
	end if
	RS1.MoveNext
next
RS1.Close
%>
</tr>
</table>
<%
end if

set RS1 = nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->