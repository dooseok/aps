<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim SQL
dim RS1

dim PO_Code
dim PO_Date
dim PO_Due_Date
dim PO_State

dim temp
dim strError
dim URL_Prev
dim URL_Next

dim strDelete

rem 객체선언
Set RS1		= Server.CreateObject("ADODB.RecordSet")

URL_Prev	= Request("URL_Prev")
URL_Next	= Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	PO_Code			= Request("PO_Code")
	PO_Date			= Request("PO_Date")
	PO_Due_Date		= Request("PO_Due_Date")
	PO_State		= Request("PO_State")

	rem DB 업데이트
	SQL = "select * from tbParts_Order where PO_Code = '"&PO_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1
		.Fields("PO_Date")			= PO_Date
		.Fields("PO_Due_Date")		= PO_Due_Date
		.Fields("PO_State")			= PO_State
		.Update
		.Close
	end with
end if

rem 객체 해제
Set RS1	= nothing
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
<form name="frmRedirect" action="<%=URL_Next%>" method=post>
<input type="hidden" name="PO_Code" value="<%=PO_Code%>">
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
<form name="frmRedirect" action="<%=URL_Prev%>" method=post>
<input type="hidden" name="PO_Code" value="<%=PO_Code%>">
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