<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
dim RS1
dim SQL

dim originB_Code
dim targetB_Code
dim B_D_No
dim newB_Version_Code

originB_Code		= Request("B_Code")
B_D_No				= Request("B_D_No")
newB_Version_Code	= Request("B_Version_Code")

set RS1 = Server.CreateObject("ADODB.RecordSet")

dim CNT1
dim CNT2
dim B_Opt_YN

CNT2 = 0
do until false
	for CNT1=0 to CNT2
		newB_Version_Code = newB_Version_Code & "(??)"
	next
	
	SQL = "select * from tbBOM where "
	SQL = SQL & " B_D_No = '"&B_D_No&"' and B_Version_Code = '"&newB_Version_Code&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		RS1.Close
		exit do
	else
		RS1.Close
	end if
	CNT2 = CNT2 + 1
loop

SQL = "select B_Opt_YN from tbBOM where B_Code="&originB_Code
RS1.Open SQL,sys_DBCon
B_Opt_YN = RS1("B_Opt_YN")
RS1.Close

SQL = "insert into tbBOM select " 
SQL = SQL & "B_D_No, "
SQL = SQL & "'"&newB_Version_Code&"', "
SQL = SQL & "B_Version_Date, "
SQL = SQL & "'N', "
SQL = SQL & "b_Issue_Date, "
SQL = SQL & "B_Class, "
SQL = SQL & "B_Tool, "
SQL = SQL & "B_Desc, "
SQL = SQL & "B_Spec, "
SQL = SQL & "'', "
SQL = SQL & "'', "
SQL = SQL & "'', "
SQL = SQL & "'', "
SQL = SQL & "B_Reg_Date, "
SQL = SQL & "B_Edit_Date, "
SQL = SQL & "B_IMD_Qty, "
SQL = SQL & "B_Point, "
SQL = SQL & "B_ST, "
SQL = SQL & "B_ST_Assm, "
SQL = SQL & "B_Standard_Time, "
SQL = SQL & "B_Standard_Time_ASM, "
SQL = SQL & "B_Tact_Time, "
SQL = SQL & "B_Tact_Time_ASM, "
SQL = SQL & "B_Memo, "
SQL = SQL & "B_IMD_MPH, "
SQL = SQL & "B_SMD_MPH, "
SQL = SQL & "B_MAN_MPH, "
SQL = SQL & "B_Check, "
SQL = SQL & "B_BuJeRyoBi, "
SQL = SQL & "'"&B_Opt_YN&"' "
SQL = SQL & "from tbBOM where B_Code = "&originB_Code
sys_DBCon.execute(SQL)

SQL = "select max(B_Code) from tbBOM where B_D_No = '"&B_D_No&"'"
RS1.open SQL,sys_DBCon
targetB_Code = RS1(0)
RS1.close


SQL = "insert into tbBOM_Sub select " 
SQL = SQL & "BS_D_No, "
SQL = SQL & targetB_Code&", "
SQL = SQL & "0, "
SQL = SQL & "0, "
SQL = SQL & "0, "
SQL = SQL & "0, "
SQL = SQL & "BS_IMD_Axial_Point, "
SQL = SQL & "BS_IMD_Radial_Point, "
SQL = SQL & "'N', "
SQL = SQL & "BS_ST, "
SQL = SQL & "BS_ST_ASM, "
SQL = SQL & "0, "
SQL = SQL & "BS_Division, "
SQL = SQL & "BS_Jeryobi, "
SQL = SQL & "BS_BarePCB "
SQL = SQL & "from tbBOM_Sub where BOM_B_Code = "&originB_Code&" order by BS_Code asc"
sys_DBCon.execute(SQL)

dim strTable
SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&originB_Code	
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

SQL = "insert into tbBOM_Qty_Archive select " 
SQL = SQL & targetB_Code&", "
SQL = SQL & "BOM_Sub_BS_D_No, "
SQL = SQL & "BOM_Sub_BS_Code, "
SQL = SQL & "Parts_P_P_No, "
SQL = SQL & "Parts_P_P_No2, "
SQL = SQL & "Parts_P_P_No2_PinYN, "
SQL = SQL & "BQ_Qty, "
SQL = SQL & "BQ_Use_YN, "
SQL = SQL & "BQ_Order, "
SQL = SQL & "BQ_Remark, "
SQL = SQL & "BQ_CheckSum, "
SQL = SQL & "BQ_P_Desc, "
SQL = SQL & "BQ_P_Spec, "
SQL = SQL & "BQ_P_Maker, "
SQL = SQL & "BOM_B_D_No "
SQL = SQL & "from "&strTable&" where BOM_B_Code = "&originB_Code&" order by BQ_Code asc"
sys_DBCon.execute(SQL)

SQL = "update tbBOM_Qty_Archive set BOM_Sub_BS_Code = (select BS_Code from tbBOM_Sub where BS_D_No = BOM_Sub_BS_D_No and BOM_B_Code = "&targetB_Code&") where BOM_B_Code = "&targetB_Code
sys_DBCon.execute(SQL)

set RS1 = nothing
%>

<html>
<head>
</head>
<body>
<form name="frmRedirect" action="b_list.asp" method="post">
<input type="hidden" name="s_bom_b_d_no" value="<%=B_D_No%>">
</form>
</body>
</html>
<script language="javascript">
alert("??? € ??????????.")
frmRedirect.submit();
</script>



<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->