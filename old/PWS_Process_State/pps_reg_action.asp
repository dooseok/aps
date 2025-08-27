<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->

<%
rem 변수선언
dim SQL
dim RS1
dim CNT1

dim PPS_Date
dim PPS_Time_Start
dim PPS_Line
dim BOM_Sub_BS_D_No
dim PPS_Qty_Required

dim temp
dim strError
dim URL_Prev
dim URL_Next

PPS_Date			= trim(Request("PPS_Date"))
PPS_Time_Start		= trim(Request("PPS_Time_Start"))
PPS_Line			= trim(Request("PPS_Line"))
BOM_Sub_BS_D_No		= ucase(trim(Request("BOM_Sub_BS_D_No")))
PPS_Qty_Required	= trim(Request("PPS_Qty_Required"))

URL_Prev			= Request("URL_Prev")
URL_Next			= Request("URL_Next")

'에러메세지가 있을 경우 실행안됨

set RS1 = Server.CreateObject("ADODB.RecordSet")

if strError = "" then
	'실적 데이터 등록
	PPS_Time_Start	= left(PPS_Time_Start,2) * 60 + right(PPS_Time_Start,2) - 500
	
	SQL = "insert into tbPWS_Process_State (PPS_Date, PPS_Time_Start,PPS_Line,BOM_Sub_BS_D_No,PPS_Qty_Required) values "
	SQL = SQL & "('"&PPS_Date&"','"&PPS_Time_Start&"','"&PPS_Line&"','"&BOM_Sub_BS_D_No&"',"&PPS_Qty_Required&")"
	sys_DBCon.execute(SQL)
end if

set RS1 = nothing
%>

<%
dim Request_Fields
dim strRequestForm
dim strRequestQueryString
for each Request_Fields in Request.Form
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
for each Request_Fields in Request.QueryString
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
if strError = "" then
%>
<form name="frmRedirect" action="pps_list.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="pps_list.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>

<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->