<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
dim RS1
dim SQL

dim PI_Code
dim PI_State

dim strError

PI_Code = request("PI_Code")

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select PI_State from tbParts_Incoming where PI_Code='"&PI_Code&"'"
RS1.Open SQL,sys_DBCon
PI_State = RS1("PI_State")
RS1.Close
set RS1 = nothing

if PI_State = "발주준비" then
	SQL = "delete from tbParts_Incoming where PI_Code='"&PI_Code&"'"
	sys_DBCon.execute(SQL)
else
	strError = strError & "*발주준비 상태인 항목만 삭제 가능합니다.\n"
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
<form name="frmRedirect" action="pi_list.asp" method=post>
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
<form name="frmRedirect" action="pi_list.asp" method=post>
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