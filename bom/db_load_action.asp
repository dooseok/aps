<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/db_header.asp" --> 

<html>
<head>
</head>
<body>
<%
dim strActionURL

if request("mode") = "print" then
	strActionURL = "b_print.asp"
else
	strActionURL = "b_model_reg_form.asp"
end if
%>
<form name="frmDB_Load" action="<%=strActionURL%>" method="post">
<input type="hidden" name="from" value="db_load_action">

<%
Dim Model_CNT
Dim Parts_CNT

Dim RS1
dim RS2
Dim SQL

Dim B_Code
dim BS_Code
dim BU_Code
dim BS_D_No

dim CNT1
Dim CNT2
Dim B_D_No

dim Parts_P_P_No 
dim BQ_Order
dim BQ_Remark
dim Parts_P_P_NO2
dim BQI_SType

dim strTable

B_Code 			= Request("B_Code")
BS_Code 		= Request("BS_Code")
BU_Code			= Request("BU_Code")
BS_D_No			= Request("BS_D_No")

if B_Code = "" then
	if BS_Code <> "" then
		B_Code = getB_Code_by_BS_Code(BS_Code)
	elseif BU_Code <> "" then
		B_Code = getB_Code_by_BU_Code(BU_Code)
	elseif BS_D_No <> "" then
		B_Code = getB_Code_by_BS_D_No(BS_D_No)
	end if
end if

dim strRow_Qty

set RS1 = Server.CreateObject("ADODB.RecordSet") 
set RS2 = Server.CreateObject("ADODB.RecordSet") 

if isnumeric(B_Code) then
	SQL  = "select top 1 B_D_No, B_Version_Current_YN from tbBOM where B_Code="&B_Code
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		response.write "요청한 도면이 DB에 없습니다. 관리자에게 문의하여주십시오."
		response.end
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable = "tbBOM_Qty"
		else
			strTable = "tbBOM_Qty_Archive"
		end if
%>
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="DNO" value="<%=RS1("B_D_No")%>">
<%
	end if
	RS1.close
else
	response.write "요청한 도면이 DB에 없습니다. 관리자에게 문의하여주십시오."
	response.end
end if

SQL = "select BS_D_No, BS_Confirm_YN, BS_Code from tbBOM_Sub where BOM_B_Code='"&B_Code&"' order by BS_D_No desc"
RS1.Open SQL,sys_DBCon
CNT2 = 1
do until RS1.Eof
	BS_D_No = RS1("BS_D_No")
%>
	<input type="hidden" name="DNOSUB" value="<%=BS_D_No%>">
	<input type="hidden" name="DNOCONFIRM" value="<%=RS1("BS_Confirm_YN")%>">
<%
	CNT2 = CNT2 + 1
	RS1.MoveNext
loop
RS1.Close
	
SQL = "select "
SQL = SQL & "BQ_Order, "
SQL = SQL & "Parts_P_P_No, "
SQL = SQL & "Parts_P_P_No2, "
SQL = SQL & "Parts_P_P_No2_PinYN, "
SQL = SQL & "BQ_P_Desc, "
SQL = SQL & "BQ_P_Spec, "
SQL = SQL & "BQ_P_Maker, "
'SQL = SQL & "P_Work_Type = (select top 1 P_Work_Type from tbParts where P_P_No = t1.Parts_P_P_No), "
SQL = SQL & "BQ_Remark, "
SQL = SQL & "BQ_CheckSum "
SQL = SQL & "from "&strTable&" t1 where BOM_B_Code="&B_Code&" and BOM_Sub_BS_D_No='"&BS_D_No&"' order by BQ_Code"
RS1.Open SQL,sys_DBCon
CNT1 = 1
Do Until RS1.Eof
	Parts_P_P_No = RS1("Parts_P_P_No")
	BQ_Order = RS1("BQ_Order")
	BQ_Remark = RS1("BQ_Remark")
	
	BQI_SType = ""
	Parts_P_P_No2 = ""
	SQL = "select top 1 BQI_SType, Parts_P_P_No2 from tbBOM_QTY_Info where Parts_P_P_No = '"&Parts_P_P_No&"'"
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
%>
	<input type="hidden" name="NO_<%=CNT1%>" value="<%=BQ_Order%>">
	<input type="hidden" name="PNO_<%=CNT1%>" value="<%=Parts_P_P_No%>">
	<input type="hidden" name="DESCRIPTION_<%=CNT1%>" value="<%=RS1("BQ_P_Desc")%>">
	<input type="hidden" name="SPEC_<%=CNT1%>" value="<%=RS1("BQ_P_Spec")%>">
	<input type="hidden" name="MAKER_<%=CNT1%>" value="<%=RS1("BQ_P_Maker")%>">
	<!--<input type="hidden" name="WORKTYPE_<%=CNT1%>" value="<$=RS1("P_Work_Type")$>">-->
	<input type="hidden" name="STYPE_<%=CNT1%>" value="<%=BQI_SType%>">
	<input type="hidden" name="REMARK_<%=CNT1%>" value="<%=BQ_Remark%>">
	<input type="hidden" name="CHECKSUM_<%=CNT1%>" value="<%=RS1("BQ_CheckSum")%>">
	<input type="hidden" name="PNO2_<%=CNT1%>" value="<%=Parts_P_P_No2%>">
	<input type="hidden" name="PNO2PinYN_<%=CNT1%>" value="<%if RS1("Parts_P_P_No2_PinYN") = "" or isnull(RS1("Parts_P_P_No2_PinYN")) then%>N<%else%>Y<%end if%>">
<%
	'이 파트넘버에 해당하는 수량 정보를 DB에서 가져옴
	SQL = "select BQ_Qty from "&strTable&" where "
	
	if BQ_Order="R" then
		SQL = SQL & "BQ_Order = 'R' and "
	else
		SQL = SQL & "BQ_Order <> 'R' and "
	end if
	
	SQL = SQL & "BOM_B_Code = "&B_Code&" and "
	SQL = SQL & "Parts_P_P_No = '"&Parts_P_P_No&"' and "
	SQL = SQL & "BQ_Remark = '"&BQ_Remark&"' "
	SQL = SQL &"order by BOM_Sub_BS_D_No desc"
	RS2.Open SQL,sys_DBCon	
	strRow_Qty = ""
	do until RS2.Eof 
		strRow_Qty = strRow_Qty&RS2("BQ_Qty")&"^"
		RS2.MoveNext
	loop
	RS2.Close
	
	if BQ_Order="R" then
%>
	<input type="hidden" name="QTY_<%=Parts_P_P_No%>_<%=BQ_Remark%>_R" value="<%=strRow_Qty%>">
<%
	else
%>
	<input type="hidden" name="QTY_<%=Parts_P_P_No%>_<%=BQ_Remark%>_X" value="<%=strRow_Qty%>">
<%
	end if
	
	CNT1 = CNT1 + 1
	
'mass--------------------------------------------
	'if strMASS_YN = "Y" and RS1("Parts_P_P_No") = "0CE4771J618" then

	'<input type="hidden" name="NO_<$=CNT1$>" value="<$=RS1("BQ_Order")$>">
	'<input type="hidden" name="PNO_<$=CNT1$>" value="0CE4776J618">
	'<input type="hidden" name="DESCRIPTION_<$=CNT1$>" value="<$=RS1("BQ_P_Desc")$>">
	'<input type="hidden" name="SPEC_<$=CNT1$>" value="SHL5.0TP35VB470M10X16 470uF 20$ 35V 753mA -40TO+85C SHL 2000HR 10X16MM 5MM STRAIGHT TP  SAMYOUNG ELECTRONICS CO., LTD.">
	'<input type="hidden" name="MAKER_<$=CNT1$>" value="삼영전자공업(주)">
	'<input type="hidden" name="WORKTYPE_<$=CNT1$>" value="<$=RS1("P_Work_Type")$>">
	'<input type="hidden" name="STYPE_<$=CNT1$>" value="<$=BQI_SType$>">
	'<input type="hidden" name="REMARK_<$=CNT1$>" value="<$=RS1("BQ_Remark")$>">
	'<input type="hidden" name="CHECKSUM_<$=CNT1$>" value="<$=RS1("BQ_CheckSum")$>">
	'<input type="hidden" name="PNO2_<$=CNT1$>" value="">
	'<input type="hidden" name="PNO2PinYN_<$=CNT1$>" value="N">-->

	'	CNT1 = CNT1 + 1
	'end if
'mass--------------------------------------------
	
	RS1.MoveNext
Loop
RS1.Close

dim BQ_Qty
dim BOM_Sub_BS_D_No
BQ_Qty = 0
Parts_P_P_No = ""
BOM_Sub_BS_D_No = ""
BQ_Remark = ""

SQL = "select "
SQL = SQL & "BOM_Sub_BS_D_No, "
SQL = SQL & "Parts_P_P_No, "
SQL = SQL & "BQ_Remark, "
SQL = SQL & "BQ_Qty, "
SQL = SQL & "BQ_Order "
SQL = SQL & "from "&strTable&" where BOM_B_Code="&B_Code
RS1.Open SQL,sys_DBCon
Do Until RS1.Eof
	BQ_Qty = RS1("BQ_Qty")
	Parts_P_P_No = RS1("Parts_P_P_No")
	BOM_Sub_BS_D_No = RS1("BOM_Sub_BS_D_No")
	BQ_Remark = RS1("BQ_Remark")
	
	if RS1("BQ_Order") = "R" then
%>
	<input type="hidden" name="QTY_<%=BOM_Sub_BS_D_No%>_<%=Parts_P_P_No%>_<%=BQ_Remark%>_R" value="<%=BQ_Qty%>">
<%
	else
%>
	<input type="hidden" name="QTY_<%=BOM_Sub_BS_D_No%>_<%=Parts_P_P_No%>_<%=BQ_Remark%>_X" value="<%=BQ_Qty%>">
<%
	end if

	RS1.MoveNext
Loop
RS1.Close


set RS2 = nothing
set RS1 = Nothing

if CNT1 = 0 then
	Parts_CNT = 30
else
	Parts_CNT = CNT1 - 1
end if

if CNT2 = 0 then
	Model_CNT = 10
else
	Model_CNT = CNT2 - 1
end if

if Model_CNT = 0 then
	Model_CNT = 1
end if
%>
<input type="hidden" name="Parts_CNT" value="<%=Parts_CNT%>">
<input type="hidden" name="Model_CNT" value="<%=Model_CNT%>">
<input type="hidden" name="oldParts_CNT" value="<%=Parts_CNT%>">
<input type="hidden" name="oldModel_CNT" value="<%=Model_CNT%>">
<input type="hidden" name="Diff_YN" value="<%=Request("Diff_YN")%>">
</form>
<script language="javascript">
<%if gM_ID="shindk" then%>
frmDB_Load.submit();
<%else%>
frmDB_Load.submit();
<%end if%>
</script>
</body>
</html>

<%
function getB_Code_by_BS_Code(BS_Code)
	dim RS1
	dim SQL
	dim B_Code
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select top 1 BOM_B_Code from tbBOM_Sub where BS_Code = '"&BS_Code&"'"
	RS1.Open SQL,sys_DBCon
	B_Code = RS1("BOM_B_Code")
	RS1.Close
	set RS1 = nothing
	getB_Code_by_BS_Code = B_Code
end function 
%>

<%
function getB_Code_by_BS_D_No(BS_D_No)
	dim RS1
	dim SQL
	dim B_Code
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select top 1 BOM_B_Code from tbBOM_Sub where "
	SQL = SQL & "BS_D_No = '"&BS_D_No&"' and "
	SQL = SQL & "exists (select top 1 B_Code from tbBOM where B_Version_Current_YN='Y' and B_Code = BOM_B_Code) "
	
	RS1.Open SQL,sys_DBCon
	
	B_Code = RS1("BOM_B_Code")
	RS1.Close
	set RS1 = nothing
	getB_Code_by_BS_D_No = B_Code
end function
%>

<%
function getB_Code_by_BU_Code(BU_Code)
	dim RS1,RS2
	dim SQL
	dim B_Code
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select top 1 BOM_B_D_No, BU_Sibang_No from tbBOM_Update_NEW where BU_Code = '"&BU_Code&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		response.write "해당되는 도면을 찾을 수 없습니다."
		response.end
	else
		SQL = "select top 1 B_Code from tbBOM where B_D_No = '"&RS1("BOM_B_D_No")&"' and B_Version_Code = '"&RS1("BU_Sibang_No")&"'"
		RS2.Open SQL,sys_DBCon
		if RS2.Eof or RS2.Bof then
			response.write "해당되는 도면을 찾을 수 없습니다."
			response.end
		else
			B_Code = RS2("B_Code")
		end if
		RS2.Close
	end if
	RS1.Close
	set RS2 = nothing
	set RS1 = nothing
	getB_Code_by_BU_Code = B_Code
end function	
%>
<!-- #include virtual = "/header/db_tail.asp" -->