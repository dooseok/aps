<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<form name='frmErrorRedirect' action='b_reg_form.asp' method=post></form>
<%
rem ��������
dim SQL
dim RS1
dim UpLoad

dim B_Code

dim NEW_YN
dim B_D_No
dim B_Version_Code
dim B_Version_Date
dim B_Version_Current_YN
dim B_Class
dim B_Tool
dim B_Desc
dim B_Spec
dim B_File_1
dim B_File_2
dim B_File_3
dim B_File_4
dim B_STate
dim B_Memo
dim B_Issue_Date
dim B_Reg_Date
dim B_Edit_Date

dim temp
dim strError
dim URL_Prev
dim URL_Next

rem ��ü����
Set RS1		= Server.CreateObject("ADODB.RecordSet")
Set UpLoad	= Server.CreateObject("Dext.FileUpLoad")

rem ���ε� �� ������ �������
UpLoad.DefaultPath = DefaultPath_BOM
UpLoad.MaxFileLen = (1024 * 1024 * 10)

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

rem ���ε� �� ���� üũ
if trim(UpLoad("B_File_1")) <> "" then
	if UpLoad("B_File_1").FileLen > UpLoad.MaxFileLen then '10�ް� �������� üũ
		strError = "����1�� 10�ް����� ���ε� �����մϴ�.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("B_File_2")) <> "" then
	if UpLoad("B_File_2").FileLen > UpLoad.MaxFileLen then '10�ް� �������� üũ
		strError = "����2�� 10�ް����� ���ε� �����մϴ�.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("B_File_3")) <> "" then
	if UpLoad("B_File_3").FileLen > UpLoad.MaxFileLen then '10�ް� �������� üũ
		strError = "����3�� 10�ް����� ���ε� �����մϴ�.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("B_File_4")) <> "" then
	if UpLoad("B_File_4").FileLen > UpLoad.MaxFileLen then '10�ް� �������� üũ
		strError = "����4�� 10�ް����� ���ε� �����մϴ�.\n"
		%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if

'SQL = "select * from tbBOM where b_d_no = '"&trim(UpLoad("B_D_No"))&"'"
'RS1.Open SQL,sys_DBCon
'if RS1.Eof or RS1.Bof then
'else
'	strError = "������ ������ �̹� ��ϵǾ��ֽ��ϴ�.\n"
'end if
'RS1.Close

rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

rem ���ε��۾�
	if trim(UpLoad("B_File_1")) <> "" then
		B_File_1 = UpLoad("B_File_1").Save(,False)
	end if
	if trim(UpLoad("B_File_2")) <> "" then
		B_File_2 = UpLoad("B_File_2").Save(,False)
	end if
	if trim(UpLoad("B_File_3")) <> "" then
		B_File_3 = UpLoad("B_File_3").Save(,False)
	end if
	if trim(UpLoad("B_File_4")) <> "" then
		B_File_4 = UpLoad("B_File_4").Save(,False)
	end if

	NEW_YN					= Trim(UpLoad("NEW_YN"))
	B_D_No					= ucase(Trim(UpLoad("B_D_No")))
	B_Version_Code			= trim(Upload("B_Version_Code"))
	B_Version_Current_YN	= Upload("B_Version_Current_YN")
	B_Version_Date			= Upload("B_Version_Date")
	B_Class					= Trim(UpLoad("B_Class"))
	B_Tool					= Trim(UpLoad("B_Tool"))
	B_Desc					= Trim(UpLoad("B_Desc"))
	B_Spec					= Trim(UpLoad("B_Spec"))
	B_File_1				= Replace(B_File_1,DefaultPath_BOM,"")
	B_File_2				= Replace(B_File_2,DefaultPath_BOM,"")
	B_File_3				= Replace(B_File_3,DefaultPath_BOM,"")
	B_File_4				= Replace(B_File_4,DefaultPath_BOM,"")
	B_Memo					= Trim(UpLoad("B_Memo"))
	B_Issue_Date			= Trim(UpLoad("B_Issue_Date"))
	B_Reg_Date				= now()
	B_Edit_Date				= now()
	
	rem DB ������Ʈ
	RS1.Open "tbBOM",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("B_D_No")			= B_D_No
		.Fields("B_Version_Code")		= B_Version_Code
		.Fields("B_Version_Current_YN")	= B_Version_Current_YN
		.Fields("B_Version_Date")		= B_Version_Date
		.Fields("B_Class")			= B_Class
		.Fields("B_Tool")			= B_Tool
		.Fields("B_Desc")			= B_Desc
		.Fields("B_Spec")			= B_Spec
		.Fields("B_File_1")			= B_File_1
		.Fields("B_File_2")			= B_File_2
		.Fields("B_File_3")			= B_File_3
		.Fields("B_File_4")			= B_File_4
		.Fields("B_Memo")			= B_Memo
		.Fields("B_Issue_Date")		= B_Issue_Date
		.Fields("B_ST")				= 500
		.Fields("B_ST_Assm")		= 500
		.Fields("B_Standard_Time")	= 0
		.Fields("B_IMD_MPH")		= 180
		.Fields("B_SMD_MPH")		= 180
		.Fields("B_MAN_MPH")		= 180
		.Fields("B_ST_Assm")		= 500
		.Fields("B_Reg_Date")		= B_Reg_Date
		.Fields("B_Edit_Date")		= B_Edit_Date
		.Fields("B_BuJeryobi")		= "0"
		.Update
		.Close
	end with
end if

SQL = "select top 1 B_Code from tbBOM where B_D_No = '"&B_D_No&"' order by B_Code desc"
RS1.Open SQL,sys_DBCon
B_Code = RS1("B_Code")
RS1.Close

if B_Version_Current_YN = "Y" then
	SQL = "update tbBOM set B_Version_Current_YN = 'N' where B_D_No = '"&B_D_No&"' and B_Code <> "&B_Code
	sys_DBCon.execute(SQL)
end if

'����� ������ ���� ������ ��, ���� �ֱٿ� ���� ���� ������ ���� �´�.
'SQL = "select top 1 B_Code from tbBOM where B_Current_YN='Y' and B_D_No='"&B_D_No&"'"
'RS1.Open SQL,sys_DBCon
'if not(RS1.Eof or RS1.Bof) then
'	call Copy_BOM(RS1("B_Code"),B_Code)	
'end if
'RS1.Close

rem ��ü ����
set UpLoad	= nothing
Set RS1		= nothing
%>


<form name="frmRedirect" action="b_list.asp" method=post>
</form>
<script language="javascript">
alert("��ϵǾ����ϴ�!");
frmRedirect.submit();
</script>


<%
function Copy_BOM(src_B_Code, dest_B_Code)
	dim RS1
	dim RS2
	dim SQL
	dim BOM_Model_BM_Code
	
	'���굵���� ī���Ѵ�.
	SQL = 		"insert into tbBOM_Sub "&vbcrlf
	SQL = SQL & "	select BS_D_No, BOM_B_Code = '"&dest_B_Code&"' from tbBOM_Sub where BOM_B_Code = '"&src_B_Code&"' order by BS_D_No asc"
	sys_DBCon.execute(SQL)
	
	'���������� ī���Ѵ�.
	SQL = 		"insert into tbBOM_Parts "&vbcrlf
	SQL = SQL & "	select BOM_B_Code = '"&dest_B_Code&"',Parts_P_P_No,BP_Order,BP_Remark,BP_Use_YN from tbBOM_Parts where BOM_B_Code = '"&src_B_Code&"' order by BP_Code asc"
	sys_DBCon.execute(SQL)
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	'���� ������ ����� ���굵���� �� ���鼭
	SQL = "select BS_Code, BS_D_No from tbBOM_Sub where BOM_B_Code = '"&src_B_Code&"' order by BS_D_No asc"
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		'��󵵹��̸鼭, ������ ���굵���� �����ڵ带 ���Ѵ�.
		SQL = "select BS_Code from tbBOM_Sub left outer join tbBOM on BOM_B_Code = B_Code where B_Code = '"&dest_B_Code&"' and BS_D_No='"&RS1("BS_D_No")&"'"
		RS2.Open SQL,sys_DBCon
		BOM_Sub_BS_Code = RS2("BS_Code")
		RS2.Close
		
		'�������̺� ���� ���굵���� �����͸� ��� ���굵������ ī���Ѵ�.
		SQL = "insert into tbBOM_Qty select BOM_Sub_BS_Code='"&BOM_Sub_BS_Code&"', BOM_Sub_BS_D_No='"&RS1("BS_D_No")&"', Parts_P_P_No, BQ_Qty, BQ_Remark, BQ_Order from tbBOM_Parts where BOM_Sub_BS_Code = '"&RS1("BS_Code")&"'"
		sys_DBCon.execute(SQL)
		
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
	set RS2 = nothing
end function
%>

<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->