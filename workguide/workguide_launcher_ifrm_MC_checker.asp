<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->


<%
dim s_PRD_Line
s_PRD_Line = request("s_PRD_Line")

dim s_IDMonitorInstance
s_IDMonitorInstance = request("s_IDMonitorInstance")
if s_IDMonitorInstance = "" then
	s_IDMonitorInstance = Request.ServerVariables("REMOTE_ADDR")
end if

if instr(application(s_PRD_Line),s_IDMonitorInstance) > 0 then '이 클라이언트로 조회한 기록이 있다면, 
	call ReloadPage(s_PRD_Line)
	response.end
else
	call MC_Check(s_PRD_Line)
	application(s_PRD_Line) = application(s_PRD_Line) & s_IDMonitorInstance & "-" 's_IDMonitorInstance를 적어놓는다.
	call ReloadPage(s_PRD_Line)
end if
%>



<%
sub MC_Check(PRD_Line)
	dim RS1
	dim SQL

	dim PRD_PartNo

	set RS1 = server.CreateObject("ADODB.RecordSet")
	if PRD_PartNo = "" then
		SQL = ""
		SQL = SQL & "select top 1 SML_PartNo from tblStatus_Monitor_Line where "
		SQL = SQL & "SML_Line='"&PRD_Line&"' and "
		SQL = SQL & "SML_Type in ('N','F','T') and "
		SQL = SQL & "SML_Process = 'START' "  
		SQL = SQL & "order by SML_Code desc "
		RS1.Open SQL,sys_DBCon
		if not(RS1.Eof or RS1.Bof) then
			PRD_PartNo = RS1("SML_PartNo")
		end if
		RS1.Close
	end if
	set RS1 = nothing
%>

<script language="javascript">
var new_PartNo = '<%=PRD_PartNo%>';
if (parent.current_PartNo != new_PartNo)
{
	parent.set_current_PartNo(new_PartNo);
}
</script>
<%
end sub
%>


<%
sub ReloadPage(PRD_Line)
%>
<script language="javascript">
function fRun()
{
	location.href="workguide_launcher_ifrm_MC_Checker.asp?s_PRD_Line=<%=PRD_Line%>&s_IDMonitorInstance=<%=s_IDMonitorInstance%>";
}

setTimeout("fRun()",1000);
</script>
<%
end sub
%>
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->