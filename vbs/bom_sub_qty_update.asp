<%
server.scripttimeout=99999
Dim sys_DBcon
Dim sys_DBconString

sys_DBConString = "Driver={SQL Server}; Server=localhost,1011; Database=msekorea; Uid=sa; Pwd=tjqltmdhs78;"

set sys_DBcon = CreateObject("adodb.connection")
sys_DBcon.ConnectionTimeout	=300
sys_DBcon.CommandTimeout 	= 300

sys_DBcon.Open sys_DBConString

Dim RS1

dim RS2
dim SQL
dim BOM_Sub_BS_D_No
dim BOM_B_D_No

dim B_IMD_Qty
dim BS_SMD_Qty
dim BS_MAN_Qty
dim BS_ASM_Qty

Set RS1 = CreateObject("ADODB.RecordSet")
Set RS2 = CreateObject("ADODB.RecordSet")

Dim dateCNT
Dim WorkDate
Dim InsertDate

for dateCNT = 1 To 10
	If dateCNT > 9 then
		WorkDate = "2009-06-" & dateCNT
	Else
		WorkDate = "2009-06-0" & dateCNT
	End If
	If dateCNT+1 > 9 then
		InsertDate = "2009-06-" & dateCNT+1
	Else
		InsertDate = "2009-06-0" & dateCNT+1
	End if
	SQL = ""
	SQL = SQL & "update " &vbcrlf
	SQL = SQL & "	tbBOM_Sub " &vbcrlf
	SQL = SQL & "set " &vbcrlf
	SQL = SQL & "	BS_IMD_Qty = BS_IMD_Qty +  " &vbcrlf
	SQL = SQL & "		isnull((select sum(PR_Amount) from tbProcess_Record  " &vbcrlf
	SQL = SQL & "		where  " &vbcrlf
	SQL = SQL & "			PR_Work_Date = '"&WorkDate&"' and " &vbcrlf
	SQL = SQL & "			PR_WorkType='�۾�' and " &vbcrlf
	SQL = SQL & "			PR_Process='IMD' and BOM_Sub_BS_D_No = BS_D_No),0), " &vbcrlf
	SQL = SQL & "	BS_MAN_Qty = BS_MAN_Qty +  " &vbcrlf
	SQL = SQL & "		isnull((select sum(PR_Amount) from tbProcess_Record  " &vbcrlf
	SQL = SQL & "		where  " &vbcrlf
	SQL = SQL & "			PR_Work_Date = '"&WorkDate&"' and " &vbcrlf
	SQL = SQL & "			PR_WorkType='�۾�' and " &vbcrlf
	SQL = SQL & "			PR_Process='MAN' and BOM_Sub_BS_D_No = BS_D_No),0), " &vbcrlf
	SQL = SQL & "	BS_SMD_Qty = BS_SMD_Qty +  " &vbcrlf
	SQL = SQL & "		isnull((select sum(PR_Amount) from tbProcess_Record  " &vbcrlf
	SQL = SQL & "		where  " &vbcrlf
	SQL = SQL & "			PR_Work_Date = '"&WorkDate&"' and " &vbcrlf
	SQL = SQL & "			PR_WorkType='�۾�' and " &vbcrlf
	SQL = SQL & "			PR_Process='SMD' and BOM_Sub_BS_D_No = BS_D_No),0), " &vbcrlf
	SQL = SQL & "	BS_ASM_Qty = BS_ASM_Qty +  " &vbcrlf
	SQL = SQL & "		isnull((select sum(PR_Amount) from tbProcess_Record  " &vbcrlf
	SQL = SQL & "		where  " &vbcrlf
	SQL = SQL & "			PR_Work_Date = '"&WorkDate&"' and " &vbcrlf
	SQL = SQL & "			PR_WorkType='�۾�' and " &vbcrlf
	SQL = SQL & "			PR_Process='ASM' and BOM_Sub_BS_D_No = BS_D_No),0) " &vbcrlf
	'response.write SQL & "<Br>"
	sys_DBCon.execute(SQL)

	SQL = "select BS_D_No from tbBOM_Sub order by bs_d_no asc"
	RS1.Open SQL,sys_DBCon
	Do until RS1.Eof
		BOM_Sub_BS_D_No	= RS1("BS_D_No")
		If Left(UCase(BOM_Sub_BS_D_No),3) = "EBR" Or Left(UCase(BOM_Sub_BS_D_No),3) = "ABQ" Or Left(UCase(BOM_Sub_BS_D_No),3) = "AEJ" Then
			BOM_B_D_No = Left(BOM_Sub_BS_D_No,9)
		Else
			BOM_B_D_No = Left(BOM_Sub_BS_D_No,10)
		End If
		
		SQL = ""
		SQL = SQL & "select b_imd_qty = (select b_imd_qty from tbBOM where b_d_no = '"&BOM_B_D_No&"' and b_current_yn='Y'), "
		SQL = SQL & " bs_smd_qty, bs_man_qty, bs_asm_qty "
		SQL = SQL & "from tbBOM_Sub where bs_d_no = '"&BOM_Sub_BS_D_No&"' "
		'response.write SQL & "<Br>"
		RS2.Open SQL,sys_DBCon
		
		B_IMD_Qty = RS2("B_IMD_Qty")
		BS_SMD_Qty = RS2("BS_SMD_Qty")
		BS_MAN_Qty = RS2("BS_MAN_Qty")
		BS_ASM_Qty = RS2("BS_ASM_Qty")
	
		If Not(isnumeric(B_IMD_Qty)) Then
			B_IMD_Qty = 0
		End If
		If Not(isnumeric(BS_SMD_Qty)) Then
			BS_SMD_Qty = 0
		End If
		If Not(isnumeric(BS_MAN_Qty)) Then
			BS_MAN_Qty = 0
		End If
		If Not(isnumeric(BS_ASM_Qty)) Then
			BS_ASM_Qty = 0
		End if
	
		SQL = ""
		SQL = SQL & "insert tbBOM_Sub_Qty_History (BOM_Sub_BS_D_No,BSQH_IMD_Qty,BSQH_SMD_Qty,BSQH_MAN_Qty,BSQH_ASM_Qty,BSQH_Date) values("
		SQL = SQL & "'"&BOM_Sub_BS_D_No&"',"
		SQL = SQL & B_IMD_Qty&","
		SQL = SQL & BS_SMD_Qty&","
		SQL = SQL & BS_MAN_Qty&","
		SQL = SQL & BS_ASM_Qty&","
		'SQL = SQL & "'"&date()&"')"
		SQL = SQL & "'"&InsertDate&"')"
	'	response.write SQL & "<Br><Br>"
		sys_DBCon.execute(SQL)
		
		RS2.Close
		
		RS1.MoveNext
	loop
	RS1.Close

next

Set RS1 = Nothing
Set RS2 = Nothing
Set sys_DBCon = Nothing
%>


