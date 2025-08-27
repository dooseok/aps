<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim strH
dim strM
dim strS

strH = hour(now())
if strH < 10 then
	strH = "0"&strH
end if
strM = minute(now())
if strM < 10 then
	strM = "0"&strM
end if
strS = second(now())
if strS < 10 then
	strS = "0"&strS
end if

dim PRD_PartNo
dim PRD_BarCode
dim PRD_Input_Date
dim PRD_Input_Time
dim PRD_Line

PRD_PartNo				= request("strCustom_Model")
PRD_BarCode				= PRD_Line&PRD_PartNo&replace(date(),"-","")&strH&strM&strS
PRD_Input_Date			= date()
PRD_Input_Time			= strH&strM
PRD_Line				= request("s_PRD_Line")


dim s_strLine
dim s_strProcess
dim s_strKeys

s_strProcess = "START"
s_strLine = PRD_Line

dim RS1
dim SQL
set RS1 = server.CreateObject("ADODB.Recordset")
SQL = "select LI_Monitor_Process from tblLine_Info where LI_Line ='"&PRD_Line&"'"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	s_strKeys = "0000000" & request("strCustom_Model") & "000000000000"
end if
RS1.Close
set RS1 = nothing


'[현황판 변경 전------------------------------
'SQL = "insert into tbPWS_Raw_Data (PRD_PartNo, PRD_BarCode, PRD_Input_Date, PRD_Input_Time, PRD_Line, PRD_Dummy_YN) values "
'SQL = SQL & "('"&PRD_PartNo&"','"&PRD_BarCode&"','"&PRD_Input_Date&"','"&PRD_Input_Time&"','"&PRD_Line&"','Y')"
'sys_DBCon.execute(SQL)
'현황판 변경 전------------------------------]
'[현황판 변경 후------------------------------
'application(PRD_Line&"_Last")=PRD_PartNo
'현황판 변경 후------------------------------]
%>

<%
if s_strLine <> "" and s_strKeys <> "" then
%>
<form name="frmRedirect" action="http://<%=gHOST%>:2080/status_monitor/barcode2db_process/barcode2db_process.asp" method=post>
<input type="hidden" name="s_strLine" value="<%=s_strLine%>">
<input type="hidden" name="s_strProcess" value="<%=s_strProcess%>">
<input type="hidden" name="s_strKeys" value="<%=s_strKeys%>">
<input type="hidden" name="s_Type" value="T">
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="about:blank" method=post>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
end if
%>
<!-- #include virtual = "/header/db_tail.asp" -->