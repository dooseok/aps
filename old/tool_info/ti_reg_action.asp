<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem ��������
dim SQL
dim RS1

dim TI_Name
dim TI_Type

dim temp
dim strError
dim URL_Prev
dim URL_Next

TI_Name	= trim(Request("TI_Name"))
TI_Type	= trim(Request("TI_Type"))

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem �����޼����� ���� ��� ����ȵ�

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select top 1 TI_Name from tbTool_Info where TI_Name='"&TI_Name&"'"
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	strError = "* ������ ���� �̹� ��ϵǾ��ֽ��ϴ�.\n"
end if
RS1.Close

if strError = "" then

	SQL = "insert into tbTool_Info (TI_Name,TI_Type) values "
	SQL = SQL & "('"&TI_Name&"','"&TI_Type&"')"
	
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
<form name="frmRedirect" action="ti_list.asp" method=post>
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
<form name="frmRedirect" action="ti_list.asp" method=post>
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