<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<% 
Response.Buffer = false 
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment;filename=" &Request("B_D_No")&".xls"
%>
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->

<%
dim SQL
dim RS1
dim RS2

dim CNT1
dim CNT2
dim CNT3
dim CNT4

dim B_Code
dim B_D_No

dim strBOM_Sub_BS_D_No
dim arrBOM_Sub_BS_D_No

dim strBS_Code
dim arrBS_Code

dim CNT_Row
dim SUM_Col

dim BQ_Qty

dim cnt_BOM_Sub
dim cnt_Parts

dim strBQ_Code
dim arrBQ_Code
dim strBQ_Qty
dim arrBQ_Qty

dim Material_Price
dim COSP_Price
dim Muldong
dim SilJeok
dim Danga
dim BuJeRyoBi

B_Code = Request("B_Code")
B_D_No = Request("B_D_No")

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

if B_D_No = "" then
	SQL = "select * from tbBOM where B_Code = '"&B_Code&"'"
	RS1.Open SQL,sys_DBCon
	B_D_No = RS1("B_D_No")
	RS1.Close
end if

SQL = "select * from tbBOM_Sub where BOM_B_Code = '"&B_Code&"' order by BS_D_No"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strBS_Code			= strBS_Code			& RS1("BS_Code")	& ";"
	strBOM_Sub_BS_D_No	= strBOM_Sub_BS_D_No	& RS1("BS_D_No")	& ";"
	RS1.MoveNext
loop
RS1.Close

arrBS_Code			= split(strBS_Code,";")
arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,";")
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	response.write replace(arrBOM_Sub_BS_D_No(CNT1),B_D_No,"") &vbtab
next
response.write vbtab
response.write vbcrlf

response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write "단가" &vbtab
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	SQL = "select top 1 BP_Price from tbBOM_Price where BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"' order by BP_Code desc"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
			Danga = 0
	else
			Danga = RS2("BP_Price")
	end if
	RS2.Close
	response.write Danga &vbtab
next
response.write vbtab
response.write vbcrlf

response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write "물동" &vbtab
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	SQL = "select sum(M_Qty) from tbMuldong where M_YYMM like '13%' and BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
			Muldong = 0
	else
			Muldong = RS2(0)
	end if
	RS2.Close
	response.write Muldong &vbtab
next
response.write vbtab
response.write vbcrlf

response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write "실적" &vbtab
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	SQL = "select sum(S_Qty) from tbSilJeok where S_YYMM like '12%' and BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
			SilJeok = 0
	else
			SilJeok = RS2(0)
	end if
	RS2.Close
	response.write SilJeok &vbtab
next
response.write vbtab
response.write vbcrlf

response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write "재료비" &vbtab
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	response.write vbtab
next
response.write vbtab
response.write vbcrlf

response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write vbtab
response.write "재료비율" &vbtab
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	response.write vbtab
next
response.write vbtab
response.write vbcrlf

response.write "No" &vbtab
response.write "P/No" &vbtab
response.write "Description" &vbtab
if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
	response.write "tbPartsMSE"&vbtab
else
	response.write "CheckSum" &vbtab
end if
response.write "Spec" &vbtab
response.write "Maker" &vbtab
response.write "Remark" &vbtab
response.write "Loc" &vbtab
if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
	response.write "tbPartsLGE"&vbtab
else
	response.write "Type" &vbtab
end if
if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
	response.write "tbMaterial"&vbtab
else
	response.write "" &vbtab
end if
if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
	response.write "COSP"&vbtab
else
	response.write "" &vbtab
end if
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	response.write "QTY" &vbtab
next
response.write "SUM" &vbtab
response.write vbcrlf 

SQL = "select distinct BS_D_No from tbBOM_Sub where BOM_B_Code = '"&B_Code&"'"
RS1.Open SQL,sys_DBCon,1
cnt_BOM_Sub = RS1.RecordCount
RS1.Close

'SQL = "select distinct Parts_P_P_No from tbBOM_Qty where BOM_Sub_BS_D_No in (select BOM_Sub_BS_D_No = BS_D_No from tbBOM_Sub where BOM_B_Code = '"&B_Code&"')"

SQL = "select * from tbBOM_Qty where BOM_Sub_BS_Code=(select min(BS_Code) from tbBOM_Sub where BOM_B_Code = '"&B_Code&"')"
RS1.Open SQL,sys_DBCon,1
cnt_Parts = RS1.RecordCount
RS1.Close

SQL = 		"select "&vbcrlf
SQL = SQL & "	BQ_Code, "&vbcrlf
SQL = SQL & "	BQ_Qty "&vbcrlf
SQL = SQL & "from tbBOM_Qty "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_Code in "&vbcrlf
SQL = SQL & "		(select BS_Code "&vbcrlf
SQL = SQL & "		from tbBOM_Sub "&vbcrlf
SQL = SQL & "		where BOM_B_Code = '"&B_Code&"') "&vbcrlf
SQL = SQL & "order by BOM_Sub_BS_D_No, BQ_Code "&vbcrlf

RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strBQ_Code	= strBQ_Code	& RS1("BQ_Code") & ";"
	strBQ_Qty	= strBQ_Qty		& RS1("BQ_Qty") & ";"
	RS1.MoveNext
loop
RS1.Close

arrBQ_Code	= split(strBQ_Code	,";")
arrBQ_Qty	= split(strBQ_Qty	,";")
%>

<%
CNT3 = 0
CNT4 = 0
for CNT1 = 1 to cnt_Parts
	SQL = 		"select "&vbcrlf
	SQL = SQL & "	Parts_P_P_No, "&vbcrlf
	SQL = SQL & "	Parts_P_P_No2, "&vbcrlf
	SQL = SQL & "	BQ_P_Desc, "&vbcrlf
	SQL = SQL & "	BQ_P_Spec, "&vbcrlf
	SQL = SQL & "	BQ_P_Maker, "&vbcrlf
	SQL = SQL & "	P_MSE_Price = (select top 1 M_Price from tbMaterial where M_P_No = Parts_P_P_No2), "&vbcrlf
	SQL = SQL & "	P_LGE_Price = (select top 1 M_Price_LGE from tbMaterial where M_P_No = Parts_P_P_No2), "&vbcrlf
	SQL = SQL & "	BQ_Qty, "&vbcrlf
	SQL = SQL & "	BQ_Remark, "&vbcrlf
	SQL = SQL & "	BQ_CheckSum, "&vbcrlf
	SQL = SQL & "	BQ_Order, "&vbcrlf
	SQL = SQL & "	P_Work_Type = (select top 1 M_Process from tbMaterial where M_P_No = Parts_P_P_No2) "&vbcrlf
	SQL = SQL & "from tbBOM_Qty "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	BQ_Code = '"&arrBQ_Code(CNT3+CNT4)&"' order by BOM_Sub_BS_D_No asc "&vbcrlf
	RS1.Open SQL,sys_DBCon
	response.write CNT_Row + CNT1 &vbtab
	response.write RS1("P_P_No") &vbtab
	response.write RS1("BQ_P_Desc") &vbtab
	if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
		response.write RS1("P_MSE_Price") &vbtab
	else
		response.write RS1("BQ_CheckSum") &vbtab
	end if
	response.write RS1("BQ_c") &vbtab
	response.write RS1("BQ_P_Maker") &vbtab
	response.write RS1("BQ_Remark") &vbtab
	response.write RS1("BQ_Order") &vbtab
	if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
		response.write RS1("P_LGE_Price") &vbtab
	else
		response.write RS1("P_Work_Type") &vbtab
	end if
	SQL = "select M_Price = M_Price/M_Package_Unit from tbMaterial where M_P_No = '"&RS1("Parts_P_P_No2")&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
		Material_Price = 0
	elseif isnumeric(RS2(0)) then
		Material_Price = RS2("M_Price")
	else
		Material_Price = 0
	end if
	RS2.Close
	if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
		response.write Material_Price &vbtab
	else
		response.write "" &vbtab
	end if
	SQL = "select CP_Price from tbCOSP_Price where Material_M_P_No = '"&RS1("Parts_P_P_No2")&"'"
	RS2.Open SQL,sys_DBCon
	if RS2.Eof or RS2.Bof then
		COSP_Price = 0
	elseif isnumeric(RS2(0)) then
		COSP_Price = RS2("CP_Price")
	else
		COSP_Price = 0
	end if
	RS2.Close
	if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
		response.write COSP_Price &vbtab
	else
		response.write "" &vbtab
	end if
	RS1.Close
	
	for CNT2 = 1 to cnt_BOM_Sub
	
		BQ_Qty = arrBQ_Qty(CNT3+CNT4)
		if isnumeric(arrBQ_Qty(CNT3+CNT4)) then
			if arrBQ_Qty(CNT3+CNT4) = 0 then
				BQ_Qty = ""
			else
				SUM_Col = SUM_Col + BQ_Qty
			end if
		else
			BQ_Qty = ""
		end if

		if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
			if BQ_Qty="" then
				BQ_Qty = "0"
			end if
		end if
		response.write BQ_Qty &vbtab
		CNT3 = CNT3 + cnt_Parts
	next
	response.write SUM_Col &vbtab
	SUM_Col = 0
	response.write vbcrlf

	CNT3 = 0
	CNT4 = CNT4 + 1
next
%>

<%
set RS2 = nothing
set RS1 = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->