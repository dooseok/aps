<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem ��������
dim SQL
dim RS1

dim MT_Code
dim MT_Date
dim MT_Company
dim MT_Remark

dim temp
dim strError
dim URL_Prev
dim URL_Next

dim strDelete

rem ��ü����
Set RS1		= Server.CreateObject("ADODB.RecordSet")

URL_Prev	= Request("URL_Prev")
URL_Next	= Request("URL_Next")

rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

	MT_Code			= Request("MT_Code")
	MT_Date			= Request("MT_Date")
	MT_Company		= Request("MT_Company")
	MT_Remark		= Request("MT_Remark")

	rem DB ������Ʈ
	SQL = "select * from tbMaterial_Transaction where MT_Code = '"&MT_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1
		.Fields("MT_Date")			= MT_Date
		.Fields("MT_Company")		= MT_Company
		.Fields("MT_Remark")		= MT_Remark
		.Update
		.Close
	end with
end if

rem ��ü ����
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
<input type="hidden" name="MT_Code" value="<%=MT_Code%>">
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
<input type="hidden" name="MT_Code" value="<%=MT_Code%>">
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