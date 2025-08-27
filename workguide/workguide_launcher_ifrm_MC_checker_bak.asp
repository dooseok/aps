<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim RS1
dim SQL

dim PRD_PartNo
dim s_PRD_Line

s_PRD_Line = request("s_PRD_Line")

set RS1 = server.CreateObject("ADODB.RecordSet")
SQL = ""
SQL = SQL & " select top 1 PRD_PartNo "
SQL = SQL & " from tbPWS_Raw_Data "
SQL = SQL & " where "
SQL = SQL & " 	PRD_Line='"&s_PRD_Line&"' and "
SQL = SQL & " 	PRD_Input_Date = '"&date()&"' and "
SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
SQL = SQL & " order by PRD_Input_Time_Detail desc "
RS1.Open SQL,sys_DBCon

PRD_PartNo = ""
if not(RS1.Eof or RS1.Bof) then
	PRD_PartNo = RS1("PRD_PartNo")
end if
RS1.Close

if PRD_PartNo = "" then
	SQL = ""
	SQL = SQL & " select top 1 PRD_PartNo "
	SQL = SQL & " from tbPWS_Raw_Data "
	SQL = SQL & " where "
	SQL = SQL & " 	PRD_Line='"&s_PRD_Line&"' and "
	SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
	SQL = SQL & " order by "
	SQL = SQL & " 	PRD_Input_Date desc, "
	SQL = SQL & " 	PRD_Input_Time_Detail desc "
	RS1.Open SQL,sys_DBCon
	if not(RS1.Eof or RS1.Bof) then
		PRD_PartNo = RS1("PRD_PartNo")
	end if
	RS1.Close
end if


set RS1 = nothing
'response.write PRD_PartNo & "<Br>"
%>

<script language="javascript">
var new_PartNo = '<%=PRD_PartNo%>';
if (parent.current_PartNo != new_PartNo)
{
	parent.set_current_PartNo(new_PartNo);
}


function fRun()
{
	location.href="workguide_launcher_ifrm_MC_Checker.asp?s_PRD_Line=<%=s_PRD_Line%>";	
}

setTimeout("fRun()",2000);
</script>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->