<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<% 
dim RS1
dim RS2

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

dim SQL
dim BP_Type
dim BP_Gap
SQL = "select * from tbBOM_Price order by BP_Code asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	SQL = "select top 1 * from tbBOM_Price where "
	SQL = SQL & "BP_Division		= '"&RS1("BP_Division")&"' 		and "
	SQL = SQL & "BOM_Sub_BS_D_No	= '"&RS1("BOM_Sub_BS_D_No")&"' 	and "
	SQL = SQL & "BP_Market			= '"&RS1("BP_Market")&"' 			and "
	SQL = SQL & "BP_Currency		= '"&RS1("BP_Currency")&"' and BP_Code < "&RS1("BP_Code")&" order by BP_Code desc"
	RS2.Open SQL,sys_DBCon
	
	BP_Gap = 0
	BP_Type = ""
	if RS2.Eof or RS2.Bof then
		BP_Type = "신규"
	else
		BP_Gap = RS1("BP_Price")-RS2("BP_Price")
		if BP_Gap > 0 then
			BP_Type = "인상"
		elseif BP_Gap = 0 then
			BP_Type = "기타"
		else
			BP_Type = "인하"
		end if
	end if
	RS2.Close
	SQL = "update tbBOM_Price set BP_Type = '"&BP_Type&"', BP_Gap = "&BP_Gap&" where BP_Code = "&RS1("BP_Code")
	sys_DBCon.execute(SQL)
	
	RS1.MoveNext
loop
RS1.Close

set RS1 = nothing
set RS2 = nothing
%>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
