<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
 
<%
dim Part_filePrefix
dim Part_Title
dim Part_Title_Eng
Part_filePrefix = "bpmd"
Part_Title		= "����"
Part_Title_Eng	= "Dev"

rem ��������
dim SQL
dim RS1

dim BPM_PartNo
dim BPM_StartDate
dim BPM_EndDate
dim BPM_Memo

dim temp
dim strError
dim URL_Prev
dim URL_Next

rem ��ü����
Set RS1		= Server.CreateObject("ADODB.RecordSet")

rem ���ε� �� ������ �������
URL_Prev	= Request("URL_Prev")
URL_Next	= Request("URL_Next")

rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

rem ���ε��۾�
	BPM_StartDate	= Trim(Request("BPM_StartDate"))
	BPM_EndDate		= Trim(Request("BPM_EndDate"))
	BPM_Memo		= Trim(Request("BPM_Memo"))
	BPM_PartNo		= Trim(Request("BPM_PartNo"))
	
	rem DB ������Ʈ
	RS1.Open "tbBOM_PartsOutSheet_Memo_"&Part_Title_Eng,sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("BPM_PartNo")	= BPM_PartNo
		.Fields("BPM_StartDate")= BPM_StartDate
		.Fields("BPM_EndDate")	= BPM_EndDate
		.Fields("BPM_Memo")		= BPM_Memo
		
		.Update
		.Close
	end with
end if

rem ��ü ����
Set RS1		= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="<%=URL_Next%>" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="<%=URL_Prev%>" method=post>

</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>

<!-- #include Virtual = "/header/db_tail.asp" -->