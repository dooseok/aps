<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim Part_filePrefix
dim Part_Title
dim Part_Title_Eng
Part_filePrefix = "bpmq"
Part_Title		= "ǰ��"
Part_Title_Eng	= "QA"

rem ��������
dim SQL
dim RS1

dim BPM_Code
dim BPM_PartNo
dim BPM_StartDate
dim BPM_EndDate
dim BPM_Memo

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

	BPM_Code		= Request("BPM_Code")
	BPM_PartNo		= Request("BPM_PartNo")
	BPM_StartDate	= Request("BPM_StartDate")
	BPM_EndDate		= Request("BPM_EndDate")
	BPM_Memo		= Request("BPM_Memo")
	
	rem DB ������Ʈ
	SQL = "select * from tbBOM_PartsOutSheet_Memo_"&Part_Title_Eng&" where BPM_Code = '"&BPM_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1
		.Fields("BPM_PartNo")		= BPM_PartNo
		.Fields("BPM_StartDate")	= BPM_StartDate
		.Fields("BPM_EndDate")		= BPM_EndDate
		.Fields("BPM_Memo")			= BPM_Memo
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
<input type='hidden' name='BPM_Code' value='<%=BPM_Code%>'>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("�����Ϸ� �Ǿ����ϴ�.");
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