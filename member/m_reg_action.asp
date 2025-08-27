<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim M_ID
dim M_Channel
dim M_Password
dim M_Part
dim M_Position
dim M_Name
dim M_Email_1
dim M_Email_2
dim M_HP
dim M_Enter_Date
dim M_Retire_Date
dim M_Authority
dim M_Use_YN

dim temp
dim strError
dim URL_Prev
dim URL_Next

M_ID			= trim(Request("M_ID"))
M_Channel		= trim(Request("M_Channel"))
M_Password		= trim(Request("M_Password"))
M_Part			= trim(Request("M_Part"))
M_Position		= trim(Request("M_Position"))
M_Name			= trim(Request("M_Name"))
M_Email_1		= trim(Request("M_Email_1"))
M_Email_2		= trim(Request("M_Email_2"))
M_HP			= trim(Request("M_HP"))
M_Enter_Date	= trim(Request("M_Enter_Date"))
M_Retire_Date	= trim(Request("M_Retire_Date"))
M_Authority		= trim(Request("M_Authority"))
M_Use_YN		= trim(Request("M_Use_YN"))
if M_Enter_Date = "" then
	M_Enter_Date	= date()
end if
URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select top 1 M_Code from tbMember where M_ID='"&M_ID&"'"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
		strError = strError & "* 동일한 아이디의 사원이 이미 등록되어있습니다.\n"
end if
RS1.Close
set RS1 = nothing
		
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	SQL = "insert into tbMember (M_ID,M_Channel,M_Password,M_Part,M_Position,M_Name,M_Email_1,M_Email_2,M_HP,M_Enter_Date,M_Authority,M_Use_YN) values "
	SQL = SQL & "('"&M_ID&"','"&M_Channel&"','"&M_Password&"','"&M_Part&"','"&M_Position&"','"&M_Name&"','"&M_Email_1&"','"&M_Email_2&"','"&M_HP&"','"&M_Enter_Date&"','"&M_Authority&"','"&M_Use_YN&"')"
	
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
<form name="frmRedirect" action="m_list.asp" method=post>

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
<form name="frmRedirect" action="m_list.asp" method=post>

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