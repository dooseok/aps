<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim SQL
dim RS1
dim RS2
dim RS3
dim RS4
dim RS5

dim arrKey

dim strTable

set RS1 = server.CreateObject("ADODB.RecordSet")
set RS2 = server.CreateObject("ADODB.RecordSet")
set RS3 = server.CreateObject("ADODB.RecordSet")
set RS4 = server.CreateObject("ADODB.RecordSet")
set RS5 = server.CreateObject("ADODB.RecordSet")

'벌크업로드 대상 B_Code 조회
SQL = "select B_Code from tbBOM where B_Version_Code = 'devtemp'"
'SQL = "select B_Code from tbBOM where B_Code = 1508"
RS1.Open SQL,sys_DBCon

do until RS1.Eof
	
	'벌크업로드 대상 서브파트넘버 조회
	SQL = "select BS_D_No,BS_Code from tbBOM_Sub where BOM_B_Code = "&RS1("B_Code")
	RS2.Open SQL,sys_DBCon
	
	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&RS1("B_Code")	
	RS3.Open SQL,sys_DBCon
	if RS3.Eof or RS3.Bof then
	else
		if RS3("B_Version_Current_YN") = "Y" then
			strTable = "tbBOM_Qty"
		else
			strTable = "tbBOM_Qty_Archive"
		end if
	end if
	RS3.Close
	
	do until RS2.Eof
		SQL = "select distinct(Parts_P_P_No+'_'+BQ_Remark) from "&strTable&" where BOM_Sub_BS_D_No <> '"&RS2("BS_D_No")&"' and BOM_B_Code = "&RS1("B_Code")
		
		RS3.Open SQL,sys_DBCon
		do until RS3.Eof 
			
			arrKey = split(RS3(0),"_")
			
			SQL = "select top 1 * from "&strTable&" where BOM_Sub_BS_D_No = '"&RS2("BS_D_No")&"' and BOM_B_Code = "&RS1("B_Code")&" and Parts_P_P_No='"&arrKey(0)&"' and BQ_Remark = '"&arrKey(1)&"'"
			
			RS4.Open SQL,sys_DBCon
			if RS4.Eof or RS4.Bof then
				SQL = "select * from "&strTable&" where BOM_B_Code = "&RS1("B_Code")&" and Parts_P_P_No='"&arrKey(0)&"' and BQ_Remark = '"&arrKey(1)&"'"
				RS5.Open SQL,sys_DBCon
				
				if not(RS5.Eof or RS5.Bof) then
					SQL = "insert into "&strTable&" (BOM_Sub_BS_D_No, BOM_Sub_BS_Code, BOM_B_Code, Parts_P_P_No, BQ_Order,BQ_P_Desc,BQ_P_Spec,BQ_Qty,BQ_Remark,BQ_P_Maker) values ("
					SQL = SQL & "'"&RS2("BS_D_No")&"',"
					SQL = SQL & RS2("BS_Code")&","
					SQL = SQL & RS1("B_Code")&","
					SQL = SQL & "'"&RS5("Parts_P_P_No")&"',"
					SQL = SQL & "'"&RS5("BQ_Order")&"',"
					SQL = SQL & "'"&replace(RS5("BQ_P_Desc"),"'","''")&"',"
					SQL = SQL & "'"&replace(RS5("BQ_P_Spec"),"'","''")&"',"
					SQL = SQL & "0,"
					SQL = SQL & "'"&RS5("BQ_Remark")&"',"
					SQL = SQL & "'"&RS5("BQ_P_Maker")&"')"
					
					response.write SQL &"<Br>"
				end if
				RS5.Close
			end if
			RS4.Close
			
			RS3.MoveNext
		loop
		RS3.Close
		
		RS2.MoveNext
	loop
	RS2.Close
	RS1.MoveNext
loop
RS1.Close

set RS4 = nothing
set RS3 = nothing
set RS2 = nothing
set RS1 = nothing
%>

<script language="javascript">
alert("Success!");
</script>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
