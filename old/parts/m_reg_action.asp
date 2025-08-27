<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim M_P_No
dim M_Spec
dim M_Desc
dim M_Additional_Info
dim M_Qty
dim M_Process

dim temp
dim strError
dim URL_Prev
dim URL_Next

M_P_No				= trim(Request("M_P_No"))
M_Spec				= trim(Request("M_Spec"))
M_Desc				= trim(Request("M_Desc"))
M_Additional_Info	= trim(Request("M_Additional_Info"))
M_Qty				= trim(Request("M_Qty"))
M_Process			= trim(Request("M_Process"))

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select top 1 M_P_No from tbMaterial where M_P_No='"&M_P_No&"'"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	strError = "* 동일한 파트넘버의 자재가 이미 등록되어있습니다.\n"
end if
RS1.Close

if strError = "" then

	if isnumeric(M_Qty) then
	else
		M_Qty = 0
	end if
	SQL = "insert into tbMaterial (M_P_No,M_Spec,M_Desc,M_Additional_Info,M_Qty,M_Process,M_Price) values "
	SQL = SQL & "('"&M_P_No&"','"&M_Spec&"','"&M_Desc&"','"&M_Additional_Info&"',"&M_Qty&",'"&M_Process&"',0)"
	response.write SQL
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
<form name="frmRedirect" action="M_list.asp" method=post>

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
<form name="frmRedirect" action="M_list.asp" method=post>

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