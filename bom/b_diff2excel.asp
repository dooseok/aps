<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<% 
dim RS1
dim CNT1
dim SQL
dim SQL2

dim strPNO
dim arrPNO

dim Column


strPNO	= trim(Request("PNO"))
arrPNO	= split(strPNO,",")
strPNO = ""
for CNT1 = 0 to ubound(arrPNO)
	if trim(arrPNO(CNT1)) <> "" then
		strPNO = strPNO & trim(arrPNO(CNT1)) &"|"
	end if
next
strPNO = left(strPNO,len(strPNO)-1)
arrPNO = split(strPNO,"|")


set RS1 = Server.CreateObject("ADODB.RecordSet")




dim B_Code
dim BS_D_No
dim B_Version_Code
dim B_Version_Current_YN

SQL = ""
SQL = SQL & "select "&vbcrlf
SQL = SQL & "	distinct t1.Parts_P_P_No, "&vbcrlf
for CNT1 = 0 to ubound(arrPNO)
	BS_D_No = left(arrPNO(CNT1),instr(arrPNO(CNT1),"-")-1)
	B_Version_Code = right(arrPNO(CNT1),len(arrPNO(CNT1))-instr(arrPNO(CNT1),"-"))
	SQL2 = "select B_Code, B_Version_Current_YN from tbBOM where B_Version_Code = '"&B_Version_Code&"' and B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No='"&BS_D_No&"')"
	RS1.Open SQL2,sys_DBCon
	B_Code = RS1("B_Code")
	B_Version_Current_YN = RS1("B_Version_Current_YN")
	RS1.Close
	
	SQL = SQL & "	["&BS_D_No&"-"&B_Version_Code&"] = isnull((select sum(t2.BQ_Qty) from "
	if B_Version_Current_YN = "Y" then
		SQL = SQL & " tbBOM_Qty t2 "
	else
		SQL = SQL & " tbBOM_Qty_Archive t2 "
	end if
	SQL = SQL & "Where t2.BOM_Sub_BS_D_No = '"&BS_D_No&"' and BOM_B_Code = "&B_Code&" and Parts_P_P_No = t1.Parts_P_P_No),0), "&vbcrlf
next
SQL = SQL & "	Diff = case "&vbcrlf
SQL = SQL & "		WHEN 0 = "&vbcrlf
SQL = SQL & "			(select max(Qty) from "&vbcrlf
SQL = SQL & "				(values "&vbcrlf
for CNT1 = 0 to ubound(arrPNO)
	BS_D_No = left(arrPNO(CNT1),instr(arrPNO(CNT1),"-")-1)
	B_Version_Code = right(arrPNO(CNT1),len(arrPNO(CNT1))-instr(arrPNO(CNT1),"-"))
	SQL2 = "select B_Code, B_Version_Current_YN from tbBOM where B_Version_Code = '"&B_Version_Code&"' and B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No='"&BS_D_No&"')"
	RS1.Open SQL2,sys_DBCon
	B_Code = RS1("B_Code")
	B_Version_Current_YN = RS1("B_Version_Current_YN")
	RS1.Close
	
	SQL = SQL & "				(isnull((select sum(t2.BQ_Qty) from "
	if B_Version_Current_YN = "Y" then
		SQL = SQL & " tbBOM_Qty t2 "
	else
		SQL = SQL & " tbBOM_Qty_Archive t2 "
	end if
	SQL = SQL & "				where t2.BOM_Sub_BS_D_No = '"&BS_D_No&"' and BOM_B_Code = "&B_Code&" and Parts_P_P_No = t1.Parts_P_P_No),0)) "
	
	if CNT1 < ubound(arrPNO) then
		SQL = SQL & ", "&vbcrlf
	else
		SQL = SQL & ") as AllQty(Qty)) "&vbcrlf
	end if
next
SQL = SQL & "				- "&vbcrlf
SQL = SQL & "			(select min(Qty) from "&vbcrlf
SQL = SQL & "				(values "&vbcrlf
for CNT1 = 0 to ubound(arrPNO)
	BS_D_No = left(arrPNO(CNT1),instr(arrPNO(CNT1),"-")-1)
	B_Version_Code = right(arrPNO(CNT1),len(arrPNO(CNT1))-instr(arrPNO(CNT1),"-"))
	SQL2 = "select B_Code, B_Version_Current_YN from tbBOM where B_Version_Code = '"&B_Version_Code&"' and B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No='"&BS_D_No&"')"
	RS1.Open SQL2,sys_DBCon
	B_Code = RS1("B_Code")
	B_Version_Current_YN = RS1("B_Version_Current_YN")
	RS1.Close

	SQL = SQL & "				(isnull((select sum(t2.BQ_Qty) from "
	if B_Version_Current_YN = "Y" then
		SQL = SQL & " tbBOM_Qty t2 "
	else
		SQL = SQL & " tbBOM_Qty_Archive t2 "
	end if
	SQL = SQL & "where t2.BOM_Sub_BS_D_No = '"&BS_D_No&"' and BOM_B_Code = "&B_Code&" and Parts_P_P_No = t1.Parts_P_P_No),0)) "
	
	if CNT1 < ubound(arrPNO) then
		SQL = SQL & ", "&vbcrlf
	else
		SQL = SQL & ") as AllQty(Qty)) "&vbcrlf
	end if
next
SQL = SQL & "		then "&vbcrlf
SQL = SQL & "			'same' "&vbcrlf
SQL = SQL & "		else "&vbcrlf
SQL = SQL & "			'!!!!' "&vbcrlf
SQL = SQL & "		end, "&vbcrlf
'160601
SQL = SQL & "	BQ_P_Desc, "&vbcrlf
SQL = SQL & "	BQ_P_Spec, "&vbcrlf
SQL = SQL & "	BQ_Remark, "&vbcrlf
SQL = SQL & "	Work_Type = (select top 1 P_Work_Type from tbParts where P_P_No = t1.Parts_P_P_No) "&vbcrlf
'SQL = SQL & "	P_Desc = (select top 1 P_Desc from tbParts where P_P_No = t1.Parts_P_P_No), "&vbcrlf
'SQL = SQL & "	P_Spec = (select top 1 P_Spec from tbParts where P_P_No = t1.Parts_P_P_No) "&vbcrlf
SQL = SQL & "from "&vbcrlf
SQL = SQL & "	tbBOM_Qty t1 "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	BOM_Sub_BS_D_No+'_'+convert(varchar(50),BOM_B_Code) in "&vbcrlf
SQL = SQL & "	("&vbcrlf
for CNT1 = 0 to ubound(arrPNO)

	BS_D_No = left(arrPNO(CNT1),instr(arrPNO(CNT1),"-")-1)
	B_Version_Code = right(arrPNO(CNT1),len(arrPNO(CNT1))-instr(arrPNO(CNT1),"-"))
	
	SQL2 = "select B_Code from tbBOM where B_Version_Code = '"&B_Version_Code&"' and B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No='"&BS_D_No&"')"
	RS1.Open SQL2,sys_DBCon
	B_Code = RS1("B_Code")
	RS1.Close

	SQL = SQL & "	'"&BS_D_No&"_"&B_Code&"'"
	if CNT1 < ubound(arrPNO) then
		SQL = SQL & ", "&vbcrlf
	else
		SQL = SQL & ") "&vbcrlf
	end if
next
SQL = SQL & "order by Parts_P_P_No asc"

RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else 
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition","attachment;filename=BOM_Diff_"&arrPNO(0)&"_.xls"
	for each Column in RS1.Fields
		response.write Column.Name
		response.write vbtab
	next
	response.write vbcrlf
	
	do until RS1.Eof
		for each Column in RS1.Fields
		    response.write Column.value
			response.write vbtab
		next
		response.write vbcrlf
		
		RS1.MoveNext
	loop
	RS1.Close
end if
set RS1 = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->