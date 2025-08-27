<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<% 
Response.Buffer = true 
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
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
dim BS_D_No

dim strBS_Code
dim arrBS_Code
dim strBOM_Sub_BS_D_No
dim arrBOM_Sub_BS_D_No

dim	strP_P_No
dim strBQ_P_Desc
dim strBQ_CheckSum
dim strBQ_P_Spec
dim strBQ_P_Maker
dim strBQ_Remark
dim strBQ_Order
dim strP_Work_Type
dim strP_P_No2
dim strP_P_No2_PinYN
dim strQty
dim strSType

dim	arrP_P_No
dim arrBQ_P_Desc
dim arrBQ_CheckSum
dim arrBQ_P_Spec
dim arrBQ_P_Maker
dim arrBQ_Remark
dim arrBQ_Order
dim arrP_Work_Type
dim arrP_P_No2
dim arrP_P_No2_PinYN
dim arrQty
dim arrSType

dim strBM_WType
dim strBM_Maker


dim cntParts
dim Parts_P_P_No
dim Parts_P_P_No2


dim M_Price
dim M_Price_LGE

dim strTable

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
Response.AddHeader "Content-Disposition","attachment;filename=" &B_D_No&".xls"
SQL = "select * from tbBOM_Sub where BOM_B_Code = '"&B_Code&"' order by BS_D_No"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	BS_D_No = RS1("BS_D_No")
	strBS_Code			= strBS_Code			& RS1("BS_Code")	& ";"
	strBOM_Sub_BS_D_No	= strBOM_Sub_BS_D_No	& BS_D_No	& ";"
	RS1.MoveNext
loop
RS1.Close

SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code	
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

SQL = "select count(BQ_Code) from "&strTable&" where BOM_B_Code = '"&B_Code&"' and BOM_Sub_BS_D_No = '"&BS_D_No&"'"
RS1.Open SQL,sys_DBCon
cntParts = RS1(0)
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
response.write vbtab
if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
	response.write vbtab
	response.write vbtab
	response.write vbtab
end if
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	response.write arrBOM_Sub_BS_D_No(CNT1) &vbtab
next
response.write vbcrlf

response.write "P/No" &vbtab
response.write "Description" &vbtab
response.write "CheckSum" &vbtab
response.write "Spec" &vbtab
response.write "LG Maker" &vbtab
response.write "Maker" &vbtab
response.write "Remark" &vbtab
response.write "Loc" &vbtab
response.write "Type" &vbtab
response.write "SType" &vbtab
response.write "P/No2" &vbtab
response.write "P/No2PinYN" &vbtab
if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
	response.write "MSE단가(M)" &vbtab
	response.write "LGE단가" &vbtab
	response.write "사급가" &vbtab
end if
for CNT1=0 to ubound(arrBOM_Sub_BS_D_No)-1
	response.write "Qty" &vbtab
next
response.write vbcrlf 

SQL = 		"select "&vbcrlf
SQL = SQL & "	Parts_P_P_No, "&vbcrlf
SQL = SQL & "	Parts_P_P_No2, "&vbcrlf ' = isnull(Parts_P_P_No2,''), "&vbcrlf
SQL = SQL & "	Parts_P_P_No2_PinYN, "&vbcrlf ' = isnull(Parts_P_P_No2_PinYN,''), "&vbcrlf
SQL = SQL & "	BQ_P_Desc, "&vbcrlf
SQL = SQL & "	BQ_P_Spec, "&vbcrlf
SQL = SQL & "	BQ_P_Maker, "&vbcrlf
'SQL = SQL & "	P_MSE_Price = isnull((select top 1 M_Price from tbMaterial where M_P_No = Parts_P_P_No),''), "&vbcrlf
'SQL = SQL & "	P_LGE_Price = isnull((select top 1 M_Price_LGE from tbMaterial where M_P_No = Parts_P_P_No),''), "&vbcrlf
SQL = SQL & "	BQ_Remark, "&vbcrlf
SQL = SQL & "	BQ_CheckSum, "&vbcrlf
'SQL = SQL & "	P_Work_Type = isnull((select top 1 P_Work_Type from tbParts where P_P_No = t1.Parts_P_P_No),''), "&vbcrlf
SQL = SQL & "	BQ_Order "&vbcrlf
SQL = SQL & "from "&strTable&" t1 "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	BOM_B_Code="&B_Code&" and BOM_Sub_BS_D_No='"&BS_D_No&"' order by BQ_Code"&vbcrlf
RS1.Open SQL,sys_DBCon
dim BQI_SType
do until RS1.Eof
	
	BQI_SType = ""
	Parts_P_P_No2 = ""
	SQL = "select top 1 BQI_SType, Parts_P_P_No2 from tbBOM_QTY_Info where Parts_P_P_No = '"&RS1("Parts_P_P_No")&"'"
	RS2.Open SQL,sys_DBCon
	if not(RS2.Eof or RS2.Bof) then
		BQI_SType = RS2("BQI_SType")
		Parts_P_P_No2 = RS2("Parts_P_P_No2")
	end if
	RS2.Close

	'해당 파트넘버가 핀이 꽂혀있다면, PNO2는 도면DB에서 가져온다.
	if RS1("Parts_P_P_No2_PinYN") = "Y" then
		Parts_P_P_No2 = RS1("Parts_P_P_No2")
	end if
	
	strP_P_No = strP_P_No & RS1("Parts_P_P_No") & "|/|"
	strBQ_P_Desc = strBQ_P_Desc & RS1("BQ_P_Desc") & "|/|"
	strBQ_CheckSum = strBQ_CheckSum & RS1("BQ_CheckSum") & "|/|"
	strBQ_P_Spec = strBQ_P_Spec & RS1("BQ_P_Spec") & "|/|"
	strBQ_P_Maker = strBQ_P_Maker & RS1("BQ_P_Maker") & "|/|"
	strBQ_Remark = strBQ_Remark & RS1("BQ_Remark") & "|/|"
	strBQ_Order = strBQ_Order & RS1("BQ_Order") & "|/|"
	'strP_Work_Type = strP_Work_Type & RS1("P_Work_Type") & "|/|"
	strP_P_No2 = strP_P_No2 & Parts_P_P_No2 & "|/|"
	strP_P_No2_PinYN = strP_P_No2_PinYN & RS1("Parts_P_P_No2_PinYN") & "|/|"
	strSType = strSType & BQI_SType & "|/|"

	RS1.MoveNext
loop
RS1.Close

SQL = 		"select "&vbcrlf
SQL = SQL & "	BQ_Qty "&vbcrlf
SQL = SQL & "from "&strTable&" t1 "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	BOM_B_Code="&B_Code&" order by BOM_Sub_BS_D_No, BQ_Code"&vbcrlf
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strQty = strQty & RS1("BQ_Qty") & "|/|"
	RS1.MoveNext
loop
RS1.Close

arrP_P_No		= split(strP_P_No,			"|/|")
arrBQ_P_Desc	= split(strBQ_P_Desc,		"|/|")
arrBQ_CheckSum	= split(strBQ_CheckSum,		"|/|")
arrBQ_P_Spec	= split(strBQ_P_Spec,		"|/|")
arrBQ_P_Maker	= split(strBQ_P_Maker,		"|/|")
arrBQ_Remark	= split(strBQ_Remark,		"|/|")
arrBQ_Order		= split(strBQ_Order,		"|/|")
'arrP_Work_Type	= split(strP_Work_Type,		"|/|")
arrP_P_No2		= split(strP_P_No2,			"|/|")
arrP_P_No2_PinYN= split(strP_P_No2_PinYN,	"|/|")
arrQty			= split(strQty,				"|/|")
arrSType		= split(strSType,			"|/|")



for CNT1 = 0 to cntParts - 1
	strBM_WType = ""
	strBM_Maker = ""
	SQL = "select top 1 BM_WType, BM_Maker "
	SQL = SQL & "	from tblBOM_Mask "
	SQL = SQL & "	where "
	SQL = SQL & "		BOM_Parts_BP_PNO = '"&arrP_P_No(CNT1)&"' and "
	SQL = SQL & "		(BM_Filter = '_' or BM_Filter like '%"&B_D_No&"%') "
	SQL = SQL & "	order by BM_Filter desc "
	RS1.Open SQL,sys_DBCon
	if not(RS1.Eof or RS1.Bof) then
		strBM_WType = RS1("BM_WType")
		strBM_Maker = RS1("BM_Maker")
	end if
	RS1.Close


	response.write arrP_P_No(CNT1) &vbtab
	response.write arrBQ_P_Desc(CNT1) &vbtab
	response.write arrBQ_CheckSum(CNT1) &vbtab
	response.write arrBQ_P_Spec(CNT1) &vbtab
	response.write arrBQ_P_Maker(CNT1) &vbtab
	response.write strBM_Maker &vbtab
	response.write arrBQ_Remark(CNT1) &vbtab
	response.write arrBQ_Order(CNT1) &vbtab
	response.write strBM_WType &vbtab
	response.write arrSType(CNT1) &vbtab
	response.write arrP_P_No2(CNT1) &vbtab
	response.write arrP_P_No2_PinYN(CNT1) &vbtab
	if instr(admin_bom_price_viewer,"-"&gM_ID&"-") > 0 then
		if trim(arrP_P_No2(CNT1)) = "" then
			Parts_P_P_No2 = trim(arrP_P_No(CNT1))
		else
			Parts_P_P_No2 = trim(arrP_P_No2(CNT1))
		end if
		
		
		SQL = "select M_Price/M_Package_Unit, M_Price_LGE/M_Package_Unit from tbMaterial where M_P_No = '"&Parts_P_P_No2&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			M_Price = 0
			M_Price_LGE = 0
		else
			M_Price = RS1(0)
			M_Price_LGE = RS1(1)
		end if
		RS1.Close
			
		if isnumeric(M_Price) then
			response.write M_Price &vbtab
		else
			response.write 0 &vbtab
		end if
		
		if isnumeric(M_Price_LGE) then
			response.write M_Price_LGE &vbtab
		else
			response.write 0 &vbtab
		end if
		
		SQL = "select CP_Price from tbCOSP_Price where Material_M_P_No = '"&Parts_P_P_No2&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			response.write 0 &vbtab
		elseif isnumeric(RS1(0)) then
			response.write RS1("CP_Price") &vbtab
		else
			response.write 0 &vbtab
		end if
		RS1.Close
	end if
	for CNT2 = 0 to ubound(arrBOM_Sub_BS_D_No)-1
		response.write arrQty(CNT1+cntParts*CNT2) &vbtab
	next
	response.write vbcrlf
next

set RS2 = nothing
set RS1 = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->