<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
dim RS1
dim SQL

dim PO_Code
dim PO_State

dim strError

PO_Code = request("PO_Code")

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select PO_State from tbParts_Outgoing where PO_Code='"&PO_Code&"'"
RS1.Open SQL,sys_DBCon
PO_State = RS1("PO_State")
RS1.Close
set RS1 = nothing

if PO_State = "출고준비" then
	SQL = "delete from tbParts_Outgoing where PO_Code='"&PO_Code&"'"
	sys_DBCon.execute(SQL)
else
	strError = strError & "*출고준비 상태인 항목만 삭제 가능합니다.\n"
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
<form name="frmRedirect" action="po_list.asp" method=post>
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
<form name="frmRedirect" action="po_list.asp" method=post>
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
<!-- #include virtual = "/header/session_check_tail.asp" -->攀