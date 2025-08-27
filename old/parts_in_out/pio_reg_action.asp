<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim PIO_Part
dim PIO_Type
dim PIO_Date
dim PIO_State

dim temp
dim strError
dim URL_Prev
dim URL_Next

PIO_Part	= trim(Request.Cookies("Admin")("M_Part"))
PIO_Type	= trim(Request("PIO_Type"))
PIO_Date	= trim(date())
PIO_State	= "요청작성중"

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨

if strError = "" then

	SQL = "insert into tbParts_In_Out (PIO_Part, PIO_Type, PIO_Date, PIO_State) values "
	SQL = SQL & "('"&PIO_Part&"','"&PIO_Type&"','"&PIO_Date&"','"&PIO_State&"')"
	sys_DBCon.execute(SQL)
end if
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
<form name="frmRedirect" action="PIO_list.asp" method=post>
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
<form name="frmRedirect" action="PIO_list.asp" method=post>
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