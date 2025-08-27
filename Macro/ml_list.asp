<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->


<%
dim RS1
dim SQL

set RS1 = server.CreateObject("ADODB.RecordSet")
SQL = "select top 24 * from tbMacro_Log order by ML_code desc"
RS1.Open SQL,sys_DBCon

do until RS1.Eof
	response.write RS1("ML_Item") & "_" & RS1("ML_UploadDate") & "<br>"
	RS1.MoveNext
loop
RS1.Close
response.write "<br>"
SQL = "select distinct PRD_Line, PRD_Box_Date = max(PRD_Box_Date), PRD_Box_Time = max(PRD_Box_Time) from tbPWS_Raw_Data where PRD_Box_Date = '"&date()&"' group by PRD_Line"
RS1.Open SQL,sys_DBCon

do until RS1.Eof
	response.write RS1("PRD_Line") & "_" & RS1("PRD_Box_Date") & "_" & RS1("PRD_Box_Time") & "<br>"
	RS1.MoveNext
loop
RS1.Close

set RS1 = nothing
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
