<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem ��������
dim SQL
dim RS1
dim UpLoad

dim ER_Code

dim ER_Title
dim ER_Content
dim ER_Reg_Date
dim ER_Edit_Date
dim ER_File_1
dim ER_File_2
dim ER_File_3
dim Member_M_ID

dim temp
dim strError
dim URL_Prev
dim URL_Next

rem ��ü����
Set RS1		= Server.CreateObject("ADODB.RecordSet")
Set UpLoad	= Server.CreateObject("Dext.FileUpLoad")

rem ���ε� �� ������ �������
UpLoad.DefaultPath = DefaultPath_Error_Reporting

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

rem ���ε� �� ���� üũ
if trim(UpLoad("ER_File_1")) <> "" then
	if UpLoad("ER_File_1").FileLen > (1024 * 1024 * 10) then '10�ް� �������� üũ
		strError = "����1�� 10�ް����� ���ε� �����մϴ�.\n"
	end if
end if
if trim(UpLoad("ER_File_2")) <> "" then
	if UpLoad("ER_File_2").FileLen > (1024 * 1024 * 10) then '10�ް� �������� üũ
		strError = "����2�� 10�ް����� ���ε� �����մϴ�.\n"
	end if
end if
if trim(UpLoad("ER_File_3")) <> "" then
	if UpLoad("ER_File_3").FileLen > (1024 * 1024 * 10) then '10�ް� �������� üũ
		strError = "����3�� 10�ް����� ���ε� �����մϴ�.\n"
	end if
end if


rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

rem ���ε��۾�
	if trim(UpLoad("ER_File_1")) <> "" then
		ER_File_1 = UpLoad("ER_File_1").Save(,False)
	end if
	if trim(UpLoad("ER_File_2")) <> "" then
		ER_File_2 = UpLoad("ER_File_2").Save(,False)
	end if
	if trim(UpLoad("ER_File_3")) <> "" then
		ER_File_3 = UpLoad("ER_File_3").Save(,False)
	end if
	
	ER_Title			= Trim(UpLoad("ER_Title"))
	ER_Content		= Trim(UpLoad("ER_Content"))
	ER_Reg_Date		= now()
	ER_Edit_Date		= now()
	ER_File_1		= Replace(ER_File_1,DefaultPath_Error_Reporting,"")
	ER_File_2		= Replace(ER_File_2,DefaultPath_Error_Reporting,"")
	ER_File_3		= Replace(ER_File_3,DefaultPath_Error_Reporting,"")
	Member_M_ID		= gM_ID
	
	rem DB ������Ʈ
	RS1.Open "tberror_reporting",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("ER_Title")			= ER_Title
		.Fields("ER_Content")		= ER_Content
		.Fields("ER_Reg_Date")		= ER_Reg_Date
		.Fields("ER_Edit_Date")		= ER_Edit_Date
		.Fields("ER_File_1")			= ER_File_1
		.Fields("ER_File_2")			= ER_File_2
		.Fields("ER_File_3")			= ER_File_3
		.Fields("Member_M_ID")		= Member_M_ID
		
		.Update
		.Close
	end with
end if

rem ��ü ����
set UpLoad	= nothing
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