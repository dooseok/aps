<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim CR_Part
dim CR_Line
dim CR_Title
dim CR_Use1
dim CR_Use2
dim CR_Amount
dim CR_Memo
dim CR_Date

dim temp
dim strError
dim URL_Prev
dim URL_Next

CR_Part		= trim(Request("CR_Part"))
CR_Line		= trim(Request("CR_Line"))
CR_Title	= trim(Request("CR_Title"))
CR_Use1		= trim(Request("CR_Use1"))
CR_Use2		= trim(Request("CR_Use2"))
CR_Amount	= trim(Request("CR_Amount"))
CR_Memo		= trim(Request("CR_Memo"))
CR_Date		= trim(Request("CR_Date"))

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨

if strError = "" then

	if isnumeric(CR_Amount) then
	else
		CR_Amount = 0
	end if
	SQL = "insert into tbCost_Report (CR_Part,CR_Line,CR_Title,CR_Use1,CR_Use2,CR_Amount,CR_Memo,CR_Date,CR_Sign_TeamJang_YN,CR_Sign_Isa_YN,CR_Sign_SaJang_YN) values "
	SQL = SQL & "('"&CR_Part&"','"&CR_Line&"','"&CR_Title&"','"&CR_Use1&"','"&CR_Use2&"',"&CR_Amount&",'"&CR_Memo&"','"&CR_Date&"','N','N','N')"
	
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
<form name="frmRedirect" action="cr_list.asp" method=post>
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
<form name="frmRedirect" action="cr_list.asp" method=post>
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