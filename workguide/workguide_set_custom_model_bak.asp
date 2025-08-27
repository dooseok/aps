<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim SQL

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

SQL = "insert into tbPWS_Raw_Data (PRD_PartNo, PRD_BarCode, PRD_Input_Date, PRD_Input_Time, PRD_Line, PRD_Dummy_YN) values "
SQL = SQL & "('"&PRD_PartNo&"','"&PRD_BarCode&"','"&PRD_Input_Date&"','"&PRD_Input_Time&"','"&PRD_Line&"','Y')"

sys_DBCon.execute(SQL)
%>

<form name="frmRedirect" action="about:blank" method=post>
</form>
<script language="javascript">
frmRedirect.submit();
</script>

<!-- #include virtual = "/header/db_tail.asp" -->