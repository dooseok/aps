<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<% 
call usePrinter()
%>

<% 
dim CNT1
dim CNT2
dim CNT3
dim CNT4

dim SQL
dim RS1
dim RS2
dim B_Code

dim B_Opt_YN

dim Bom_Sub_Cnt
dim OverRemark
dim arrRemark
dim strRemark


dim strPD_Desc

dim ComplexKey
dim dicComplexAcc
set dicComplexACC =Server.CreateObject("Scripting.Dictionary")
dim arrtempQty

Bom_Sub_Cnt = 0

dim PrintRow_under26
dim PrintRow_over26
dim PrintRow_under26_Blank
dim PrintRow_over26_Blank

'ÆäÀÌÁö »çÀÌÁî Á¶Àý º¯¼ö-------------------------------------------------------------------------
PrintRow_under26 = 52 '¸ðµ¨ ¼ö°¡ 26°³ ÀÌÇÏÀÏ ¶§ ÇÑÆäÀÌÁöÀÇ Çà ¼ö
PrintRow_under26_Blank = 11
PrintRow_over26 = 55 '¸ðµ¨ ¼ö°¡ 26°³ ÃÊ°úÀÏ ¶§ ÇÑÆäÀÌÁöÀÇ Çà ¼ö
PrintRow_over26_Blank = 14
'----------------------------------------------------------------------------------------

B_Code = Request("B_Code")
set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")
call Material_Guide()


'Diff===
dim Diff_YN
dim Diff_Disable_YN
Diff_YN = Request("Diff_YN")

SQL = "select top 1 B_Code, Bom_Sub_Cnt  from vwB_List where B_D_No = '"&Request("DNO")&"' and B_Code < "&B_Code&" order by B_Code desc"
RS1.open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	Diff_YN = "N"
	Diff_Disable_YN = "Y"
else
	Diff_B_Code = RS1(0)
	Bom_Sub_Cnt		 = RS1("Bom_Sub_Cnt")
end if
RS1.Close

strPD_Desc = "-"
SQL = "select BMDD_Desc_BOM from tblBOM_Mask_Desc_Detail"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strPD_Desc = strPD_Desc & RS1("BMDD_Desc_BOM") & "-"
	RS1.MoveNext
loop
RS1.Close
strPD_Desc = lcase(strPD_Desc)

if Diff_YN = "Y" then
	
	dim strComplexKey
	dim strComplexMaker
	dim strComplexQty
	dim strComplexAcc
	dim arrComplexKey
	dim arrComplexMaker
	dim arrComplexQty
	dim arrComplexAcc

	dim dicParts
	dim oldQty
	dim bChangeQty
	dim bChangeMaker

	dim currentPNO
	dim currentNO
	
	dim Diff_B_Code
	
	dim strTable
	
	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&Diff_B_Code
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable = "tbBOM_Qty"
		else
			strTable = "tbBOM_Qty_Archive"
		end if
	end if
	RS1.Close
	
	SQL = 		"select "&vbcrlf
	SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
	SQL = SQL & "	Parts_P_P_No, "&vbcrlf
	SQL = SQL & "	BQ_Remark, "&vbcrlf
	SQL = SQL & "	BQ_P_Maker, "&vbcrlf
	SQL = SQL & "	BQ_Order, "&vbcrlf
	SQL = SQL & "	BQ_Qty "&vbcrlf
	SQL = SQL & "from "&strTable&" "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	BOM_B_Code="&Diff_B_Code&" order by BOM_Sub_BS_D_No, BQ_Code"&vbcrlf
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		if ucase(RS1("BQ_Order")) = "R" then
			strComplexKey	= strComplexKey & RS1("BOM_Sub_BS_D_No")&"_"&RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")& "_R/|"
		else
			strComplexKey	= strComplexKey & RS1("BOM_Sub_BS_D_No")&"_"&RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")& "_X/|"
		end if
		strComplexMaker	= strComplexMaker & RS1("BQ_P_Maker")& "/|"
		strComplexQty	= strComplexQty & RS1("BQ_Qty")& "/|"
		RS1.MoveNext
	loop
	RS1.Close
	
	arrComplexKey	= split(strComplexKey,"/|")
	arrComplexMaker	= split(strComplexMaker,"/|")
	arrComplexQty	= split(strComplexQty,"/|")
	
	for CNT1 = 0 to ubound(arrComplexKey) - 1
		CNT3 = 0
		for CNT2 = 0 to CNT1
			if cstr(arrComplexKey(CNT1)) = cstr(arrComplexKey(CNT2)) then
				CNT3 = CNT3 + 1
			end if
		next
		strComplexACC	= strComplexACC & CNT3& "/|"
	next
	arrComplexACC = split(strComplexACC,"/|")
	
	set dicParts =Server.CreateObject("Scripting.Dictionary")
	
	SQL = 		"select "&vbcrlf
	SQL = SQL & "	Parts_P_P_No, "&vbcrlf
	SQL = SQL & "	BQ_Order, "&vbcrlf
	SQL = SQL & "	BQ_Remark "&vbcrlf
	SQL = SQL & "from "&strTable&" "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	BOM_B_Code="&Diff_B_Code
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		if ucase(RS1("BQ_Order")) = "R" then
			if not(dicParts.Exists(RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_R")) then
				dicParts.Add RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_R",0
			end if
		else
			if not(dicParts.Exists(RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_X")) then
				dicParts.Add RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_X",0
			end if
		end if
		RS1.MoveNext
	loop
	RS1.Close
end if
'Diff===




dim oldModel_CNT
dim oldParts_CNT

dim Model_CNT
dim Parts_CNT

Model_CNT		= Request.Form("Model_CNT")
if trim(Model_CNT) = "" then
	Model_CNT	= 10
end if
oldModel_CNT	= Request.Form("oldModel_CNT")

Parts_CNT		= Request.Form("Parts_CNT")
if trim(Parts_CNT) = "" then
	Parts_CNT	= 30
end if
oldParts_CNT	= Request.Form("oldParts_CNT")

dim COL

dim div_top
dim div_left

div_top		= 60
div_left	= 10

dim arrMODEL
dim MODEL
dim arrDNOSUB
dim DNOSUB
dim arrQTY
dim QTY
dim strQty

'dim strBM_WType
dim strBM_Desc
dim strBM_Maker

dim PageSize
%>


<%
COL = 1
%>

<%
dim B_Version_Code
dim DNO

DNO = Request("DNO")

SQL = "select B_Code,B_Version_Code, B_Opt_YN from tbBOM where B_D_No = '"&Request("DNO")&"' order by B_Code desc"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	B_Version_Code = "" 
else
	B_Version_Code = RS1("B_Version_Code")
end if
RS1.Close

if Model_CNT <= 26 then
%>
<table width="<%=25*26+50+100+55+195+600+150+150%>px" cellpadding=0 cellspacing=0 border=0 style="border-collapse:collapse; font-size:8pt; font-family:µ¸¿ò; text-align:center; table-layout:fixed">
<%
elseif  Model_CNT <= 30 then
%>
<table width="<%=25*30+50+100+55+195+600+150+150%>px" cellpadding=0 cellspacing=0 border=0 style="border-collapse:collapse; font-size:8pt; font-family:µ¸¿ò; text-align:center; table-layout:fixed">
<%
end if

arrDNOSUB		= split(Request("DNOSUB"),", ")

'DNOSUB Á¤·Ä
dim slDNOSUB
dim strDNOSUB

set slDNOSUB = server.createObject("System.Collections.Sortedlist")

for CNT1 = 0 to ubound(arrDNOSUB)
	if arrDNOSUB(CNT1) <> "" then
		slDNOSUB.add arrDNOSUB(CNT1), ""
	end if
next
for CNT1 = slDNOSUB.count - 1 to 0 step -1
  strDNOSUB = strDNOSUB & slDNOSUB.getKey(CNT1) & ","
next
arrDNOSUB = split(strDNOSUB,",")
set slDNOSUB = nothing

call Header()

COL = 1


PageSize = 0
for CNT1 = 1 to Parts_CNT
	'strBM_WType = ""
	strBM_Desc = ""
	strBM_Maker = ""
	'SQL = "select top 1 BM_WType, BM_Desc, BM_Maker "
	SQL = "select top 1 BM_Desc, BM_Maker "
	SQL = SQL & "	from tblBOM_Mask "
	SQL = SQL & "	where "
	SQL = SQL & "		BOM_Parts_BP_PNO = '"&Request("PNO_"+CSTR(CNT1))&"' and "
	SQL = SQL & "		(BM_Filter = '_' or BM_Filter like '%"&Request("DNO")&"%') "
	SQL = SQL & "	order by BM_Filter desc "
	RS2.Open SQL,sys_DBCon
	if not(RS2.Eof or RS2.Bof) then
		'strBM_WType = RS2("BM_WType")
		strBM_Desc = RS2("BM_Desc")
		strBM_Maker = RS2("BM_Maker")
	end if
	RS2.Close

'ÆäÀÌÁöº° Header ½ÃÀÛ----------------------------------------------------------------------
	if Model_CNT <= 26 then
		if PageSize > PrintRow_under26 and PageSize mod (PrintRow_under26+1) = 0 then
%>
</table>
<%for CNT2 = 1 to PrintRow_under26_Blank%><img src="/img/blank.jpg" width=1px height=1px><br><%next%>
<table width="<%=25*26+50+100+55+195+600+150+150%>px" cellpadding=0 cellspacing=0 border=0 style="border-collapse:collapse; font-size:8pt; font-family:µ¸¿ò; text-align:center; table-layout:fixed">
<%
			call Header()
		end if
	elseif  Model_CNT <= 30 then
		if PageSize > PrintRow_over26 and PageSize mod (PrintRow_over26+1) = 0 then
%>
</table>
<%for CNT2 = 1 to PrintRow_over26_Blank%><img src="/img/blank.jpg" width=1px height=1px><br><%next%>
<table width="<%=25*30+50+100+55+195+600+150+150%>px" cellpadding=0 cellspacing=0 border=0 style="border-collapse:collapse; font-size:8pt; font-family:µ¸¿ò; text-align:center; table-layout:fixed">
<%
			call Header()
		end if
	end if
'ÆäÀÌÁöº° Header ³¡----------------------------------------------------------------------
%>
	<tr height=22px>
<%
	strQty = 1
	arrQTY = split(strQty,", ")
	for CNT2 = 1 to Model_CNT - ubound(arrQTY)
		strQty = strQty & ", 0"
	next
	arrQTY = split(strQty,", ")
		
	if Model_CNT <= 26 then
		for CNT2 = 1 to 26-Model_CNT
%>
	<td width=25px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden">&nbsp;</div></td><%COL = COL + 1%>
<%
		next
	elseif  Model_CNT <= 30 then
		for CNT2 = 1 to 30-Model_CNT
%>
	<td width=25px align=center  style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden">&nbsp;</div></td><%COL = COL + 1%>
<%
		next
	end if
	
	for CNT2 = 1 to Model_CNT
	
		if ucase(Request("NO_"+CSTR(CNT1))) = "R" then
			ComplexKey = arrDNOSUB(CNT2-1-(Model_CNT-oldModel_CNT))&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_R"
		else
			ComplexKey = arrDNOSUB(CNT2-1-(Model_CNT-oldModel_CNT))&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_X"
		end if
	
		QTY = ""
		if strQty <> "" Then
			if CNT2 <= int(Model_CNT-oldModel_CNT) then
				QTY = ""
			else
				QTY = arrQTY(CNT2-1-(Model_CNT-oldModel_CNT))
				'if isNumeric(QTY) then
				'	if QTY <= 0 then
				'		QTY = ""
				'	end if
				'end if
			end if
		end If
		
		if strDNOSUB = "" then
			Qty = 0
		else	
			if ucase(Request("NO_"+CSTR(CNT1))) = "R" then
				Qty = Request("QTY_"&arrDNOSUB(CNT2-1-(Model_CNT-oldModel_CNT))&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_R")
			else
				Qty = Request("QTY_"&arrDNOSUB(CNT2-1-(Model_CNT-oldModel_CNT))&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_X")
			end if
		end if
		
		'¼ö·® Áßº¹ÀÎ °æ¿ì,
		if dicComplexAcc.Exists(ComplexKey) then
			dicComplexAcc.Item(ComplexKey) = cint(dicComplexAcc.Item(ComplexKey)) + 1
		else
			dicComplexAcc.Add ComplexKey,1
		end if	
		if instr(Qty,",") > 0 then
			arrtempQty = split(Qty,",")
			'response.write ComplexKey & "<br>"
			if ubound(arrtempQty) >= dicComplexAcc.Item(ComplexKey)-1 then
				Qty = trim(arrtempQty(dicComplexAcc.Item(ComplexKey)-1))
			end if
		end if
		
		bChangeQty = "N"
		bChangeMaker = "N"
		
		if isNumeric(QTY) then
			if QTY <= 0 then
				QTY = "&nbsp;"
			end if
		end if
%>	
	<td width=25px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%=QTY%></div></td><%COL = COL + 1%>
<%
	next
	
	
	arrRemark = split(replace(Request("REMARK_"+CSTR(CNT1))," ",""),",")
	
	OverRemark = 0
	if ubound(arrRemark) > 35 * 10 then
		OverRemark = 11
	elseif ubound(arrRemark) > 35 * 9 then
		OverRemark = 10
	elseif ubound(arrRemark) > 35 * 8 then
		OverRemark = 9
	elseif ubound(arrRemark) > 35 * 7 then
		OverRemark = 8
	elseif ubound(arrRemark) > 35 * 6 then
		OverRemark = 7
	elseif ubound(arrRemark) > 35 * 5 then
		OverRemark = 6
	elseif ubound(arrRemark) > 35 * 4 then
		OverRemark = 5
	elseif ubound(arrRemark) > 35 * 3 then
		OverRemark = 4
	elseif ubound(arrRemark) > 35 * 2 then
		OverRemark = 3
	elseif ubound(arrRemark) > 35 * 1 then
		OverRemark = 2
	elseif len(Request("REMARK_"+CSTR(CNT1))) > 19 then
		OverRemark = 1
	end if
		
%>
	<td width=50px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if isnull(Request("NO_"+CSTR(CNT1))) or Request("NO_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%=Request("NO_"+CSTR(CNT1))%><%end if%></div></td><%COL = COL + 1%>
	<td width=100px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if isnull(Request("PNO_"+CSTR(CNT1))) or Request("PNO_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%=Request("PNO_"+CSTR(CNT1))%><%end if%></div></td><%COL = COL + 1%>
	<!--<td width=100px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if isnull(Request("PNO2_"+CSTR(CNT1))) or Request("PNO2_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%=Request("PNO2_"+CSTR(CNT1))%><%end if%></div></td><%COL = COL + 1%>
	<td width=50px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if isnull(Request("WORKTYPE_"+CSTR(CNT1))) or Request("WORKTYPE_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%=Request("WORKTYPE_"+CSTR(CNT1))%><%end if%></div></td><%COL = COL + 1%>-->
	<td width=55px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden; letter-spacing:2px; font-family:Consolas;"><%if isnull(Request("CHECKSUM_"+CSTR(CNT1))) or Request("CHECKSUM_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%=replace(replace(Request("CHECKSUM_"+CSTR(CNT1)),"[",""),"]","")%><%end if%></div></td><%COL = COL + 1%>
	<td width=195px align=left style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if strBM_Desc = "" then%>&nbsp;<%else%><%=strBM_Desc%><%end if%></div></td><%COL = COL + 1%>
	<td width=600px align=left style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if isnull(Request("SPEC_"+CSTR(CNT1))) or Request("SPEC_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%=Request("SPEC_"+CSTR(CNT1))%><%end if%></div></td><%COL = COL + 1%>
	<td width=150px align=left style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if isnull(Request("MAKER_"+CSTR(CNT1))) or Request("MAKER_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%=strBM_Maker%><%end if%></div></td><%COL = COL + 1%>
	<td width=150px align=left style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%if isnull(Request("REMARK_"+CSTR(CNT1))) or Request("REMARK_"+CSTR(CNT1)) = "" then%>&nbsp;<%else%><%if OverRemark = 0 then%><%=Request("REMARK_"+CSTR(CNT1))%><%else%>&nbsp;<%end if%><%end if%></div></td><%COL = COL + 1%>
</tr>
<%
	COL = 1
	PageSize = PageSize + 1
	
	for CNT3 = 1 to OverRemark
	
'ÆäÀÌÁöº° Header ½ÃÀÛ----------------------------------------------------------------------	
		if Model_CNT <= 26 then
			if PageSize > PrintRow_under26 and PageSize mod (PrintRow_under26+1) = 0 then
%>
	</table>
	<%for CNT2 = 1 to PrintRow_under26_Blank%><img src="/img/blank.jpg" width=1px height=1px><br><%next%>
	<table width="<%=25*26+50+100+55+195+600+150+150%>px" cellpadding=0 cellspacing=0 border=0 style="border-collapse:collapse; font-size:8pt; font-family:µ¸¿ò; text-align:center; table-layout:fixed">
<%
				call Header()
			end if
		elseif  Model_CNT <= 30 then
			if PageSize > PrintRow_over26 and PageSize mod (PrintRow_over26+1) = 0 then
%>
	</table>
	<%for CNT2 = 1 to PrintRow_over26_Blank%><img src="/img/blank.jpg" width=1px height=1px><br><%next%>
	<table width="<%=25*30+50+100+55+195+600+150+150%>px" cellpadding=0 cellspacing=0 border=0 style="border-collapse:collapse; font-size:8pt; font-family:µ¸¿ò; text-align:center; table-layout:fixed">
<%
				call Header()
			end if
		end if
'ÆäÀÌÁöº° Header ³¡----------------------------------------------------------------------

		strRemark = ""
		for CNT4 = (35*(CNT3-1)) to 35*CNT3-1
			if CNT4 <= ubound(arrRemark) then
				strRemark = strRemark & arrRemark(CNT4) &","
			end if
		next
		
		if right(strRemark,1) = "," then
			strRemark = left(strRemark,len(strRemark)-1)
		end if
		
		strRemark = strRemark & "&nbsp;"
%>
	<tr>	
<%		
		if Model_CNT <= 26 then
			for CNT2 = 1 to 26-Model_CNT
%>
	<td width=25px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden">&nbsp;</div></td><%COL = COL + 1%>
<%
			next
		elseif  Model_CNT <= 30 then
			for CNT2 = 1 to 30-Model_CNT
%>
	<td width=25px align=center  style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden">&nbsp;</div></td><%COL = COL + 1%>
<%
			next
		end if
		for CNT2 = 1 to Model_CNT
%>
		<td width=25px align=center style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden">&nbsp;</div></td><%COL = COL + 1%>
<%		
		next
%>
		<td colspan=7 align=right style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><div style="width:100%; height:22px; overflow:hidden"><%=strRemark%></div></td><%COL = COL + 1%>
	</tr>
<%			
		PageSize = PageSize + 1
	next
next

set RS1 = nothing
set RS2 = nothing
%>
</table>
</div>


<%
sub Header
%>
<tr height=100px>
<%
dim CNT1
dim CNT2

if Model_CNT <= 26 then
	for CNT2 = 1 to 26-Model_CNT
%>
	<td width=25px align=center valign=top style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
<%
		next
elseif  Model_CNT <= 30 then
		for CNT2 = 1 to 30-Model_CNT
%>
	<td width=25px align=center valign=top style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
<%
	next
end if
	
for CNT1 = 1 to Model_CNT
	DNOSUB		= ""
	if Request("DNOSUB") <> "" then
		if CNT1 <= int(Model_CNT-oldModel_CNT) then
			DNOSUB		= ""
		else
			DNOSUB		= arrDNOSUB(CNT1-1-(Model_CNT-oldModel_CNT))
		end if
	end if
%>
	<td width=25px align=center valign=top style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;"><input class="trans_obj" type="text" name="DNOSUB" value="<%=DNOSUB%>" style="writing-mode:tb-rl; font-family:µ¸¿ò; font-size:8pt; width:12px; height:77%; text-align:center; background-color:white; border:0px"></td><%COL = COL + 1%>
<%
	COL = COL + 1
next
%>
	<td width=50px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">WORK</td><%COL = COL + 1%>
	<td width=100px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
	<!--<td width=100px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
	<td width=50px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>-->
	<td width=50px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
	<td width=200px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
	<td width=600px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
	<td width=150px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
	<td width=150px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
	
</tr>
<%
COL = 1
%>
<tr height=22px>
<%
if Model_CNT <= 26 then
	for CNT2 = 1 to 26-Model_CNT
%>
	<td width=25px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
<%
		next
elseif  Model_CNT <= 30 then
		for CNT2 = 1 to 30-Model_CNT
%>
	<td width=25px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">&nbsp;</td><%COL = COL + 1%>
<%
	next
end if
for CNT2 = 1 to Model_CNT
%>
	<td width=25px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">Q</td><%COL = COL + 1%>
<%
next
%>
	<td width=50px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">NO</td><%COL = COL + 1%>
	<td width=100px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">P/NO</td><%COL = COL + 1%>
	<!--<td width=100px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">P/No2</td><%COL = COL + 1%>
	<td width=50px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">TYPE</td><%COL = COL + 1%>-->
	<td width=50px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">C/S</td><%COL = COL + 1%>
	<td width=200px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">DESCRIPTION</td><%COL = COL + 1%>
	<td width=600px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">SPEC</td><%COL = COL + 1%>
	<td width=150px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">MAKER</td><%COL = COL + 1%>
	<td width=150px style="font-family:µ¸¿ò; border-collapse: collapse; border:0.25pt dotted darkgray;">REMARK</td><%COL = COL + 1%>
	
</tr>

<%
end Sub
%>
</table>

<script language="javascript">

function BOM_Print()
{
	//factory.printing.printer			= "Microsoft Print to PDF"; 
	factory.printing.header				= "<%=DNO%> - <%=B_Version_Code%>  page:&p/&P";
	factory.printing.footer				= "<%=now()%>";
	factory.printing.portrait			= false;
	factory.printing.leftMargin		= 1;
	factory.printing.rightMargin	= 1;
	factory.printing.topMargin		= 1;
	factory.printing.bottomMargin	= 1;
	factory.printing.print(false);
	window.close();

}

<%if Bom_Sub_Cnt > 30 then%>
alert("¿É¼ÇÀÌ 30°³¸¦ ÃÊ°úÇÏ´Â °æ¿ì¿¡´Â ÀÎ¼â¸¦ Áö¿øÇÏÁö ¾Ê½À´Ï´Ù.\nµµ¸éÀ» ¿¢¼¿·Î ´Ù¿î·ÎµåÇÏ¿© Ãâ·ÂÇØ ÁÖ½Ê½Ã¿À.");
window.close();
<%elseif gM_ID <> "shindk" then%>
setTimeout("BOM_Print()",2000);
<%end if%>
</script>
<!-- #include virtual = "/header/db_tail.asp" -->
