<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim SQL
dim RS1

dim M_Code
dim M_Password
dim M_Email_1
dim M_Email_2
dim M_HP

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

	M_Code		= Request("M_Code")
	M_Password	= Request("M_Password")
	M_Email_1	= Request("M_Email_1")
	M_Email_2	= Request("M_Email_2")
	M_HP		= Request("M_HP")
	
	rem DB 업데이트
	SQL = "select * from tbMember where M_Code = '"&M_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1
		.Fields("M_Password")	= M_Password
		.Fields("M_Email_1")	= M_Email_1
		.Fields("M_Email_2")	= M_Email_2
		.Fields("M_HP")			= M_HP
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

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("정보가 변경되었습니다.\n다시 로그인해주십시오.");
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="<%=URL_Prev%>" method=post>

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