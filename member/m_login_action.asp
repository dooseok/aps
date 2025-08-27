<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->

<%
dim RS1
dim SQL

dim Login_Type
dim login_URL

dim M_Code
dim M_Channel
dim M_ID
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

dim strError

login_URL	= trim(Request("login_URL"))
M_ID		= lcase(trim(Request("M_ID")))
M_Password	= trim(Request("M_Password"))

set RS1 = server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbMember where (M_Retire_Date is null or M_Retire_Date = '') and M_Use_YN = 'Y' and M_ID='"&M_ID&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	strError = "아이디가 존재하지 않습니다."
elseif RS1("M_Password") <> M_Password then
	strError = "비밀번호가 일치하지 않습니다."
else
	M_Code			= RS1("M_Code")
	M_Channel		= RS1("M_Channel")
	M_ID			= RS1("M_ID")
	M_Password		= RS1("M_Password")
	M_Part			= RS1("M_Part")
	M_Position		= RS1("M_Position")
	M_Name			= RS1("M_Name")
	M_Email_1		= RS1("M_Email_1")
	M_Email_2		= RS1("M_Email_2")
	M_HP			= RS1("M_HP")
	M_Enter_Date	= RS1("M_Enter_Date")
	M_Retire_Date	= RS1("M_Retire_Date")
	M_Authority		= RS1("M_Authority")
	M_Use_YN		= RS1("M_Use_YN")
end if
RS1.Close
set RS1 = nothing

if isnull(M_Email_1) then
	M_Email_1 = ""
end if
if isnull(M_Email_2) then
	M_Email_2 = ""
end if

if strError = "" then
	Response.cookies("Admin")("M_Code")			= M_Code
	Response.cookies("Admin")("M_Channel")		= M_Channel
	Response.cookies("Admin")("M_ID")			= M_ID
	Response.cookies("Admin")("M_Password")		= M_Password
	Response.cookies("Admin")("M_Part")			= M_Part
	Response.cookies("Admin")("M_Position")		= M_Position
	Response.cookies("Admin")("M_Name")			= M_Name
	Response.cookies("Admin")("M_Email_1")		= M_Email_1
	Response.cookies("Admin")("M_Email_2")		= M_Email_2
	Response.cookies("Admin")("M_HP")			= M_HP
	Response.cookies("Admin")("M_Enter_Date")	= M_Enter_Date
	if isnull(M_Retire_Date) then
		Response.cookies("Admin")("M_Retire_Date")	= ""
	else
		Response.cookies("Admin")("M_Retire_Date")	= M_Retire_Date
	end if
	Response.cookies("Admin")("M_Authority")	= M_Authority
	Response.cookies("Admin")("M_Use_YN")		= M_Use_YN
	Response.cookies("Admin").Path				= "/"
end if
%>

<!-- #include virtual = "/header/db_tail.asp" -->

<%
if strError = "" then
%>
<form name="back_form" method="post" action="/index.asp">
</form>
<script language="javascript">
back_form.submit();
</script>
<%
else
%>
<script language="javascript">
alert("<%=strError%>");
location.href='/member/m_logout_action.asp'
</script>
<%
end if
%>