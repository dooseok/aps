<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem ��������
dim SQL
dim RS1
dim UpLoad

dim N_Code
dim N_Title
dim N_Content
dim N_Edit_Date
dim N_File_1
dim N_File_2
dim N_File_3
dim Member_M_ID

dim oldN_File_1
dim oldN_File_2
dim oldN_File_3

dim temp
dim strError
dim URL_Prev
dim URL_Next

Dim strDelete

rem ��ü����
Set RS1		= Server.CreateObject("ADODB.RecordSet")
Set UpLoad	= Server.CreateObject("Dext.FileUpLoad")

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

rem ���ε� �� ������ �������
UpLoad.DefaultPath = DefaultPath_Notice

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

strDelete	= UpLoad("strDelete")

SQL = "select Member_M_ID from tbNotice where N_Code = '"&UpLoad("N_Code")&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	strError = strError & "*�ۼ��������� ȸ��DB���� ã�� �� �����ϴ�.\n*�����ڿ��� �����Ͽ� �ֽʽÿ�.\n"
else
	if lcase(RS1("Member_M_ID")) <> lcase(gM_ID) then
		strError = strError & "*�ۼ��� ������ ���� ������ ������ �� �ֽ��ϴ�.\n"
	end if
end if
RS1.Close

rem ���ε� �� ���� üũ
if trim(UpLoad("N_File_1")) <> "" then
	if UpLoad("N_File_1").FileLen > (1024 * 1024 * 10) then '10�ް� �������� üũ
		strError = "����1�� 10�ް����� ���ε� �����մϴ�.\n"
	end if
end if
if trim(UpLoad("N_File_2")) <> "" then
	if UpLoad("N_File_2").FileLen > (1024 * 1024 * 10) then '10�ް� �������� üũ
		strError = "����2�� 10�ް����� ���ε� �����մϴ�.\n"
	end if
end if
if trim(UpLoad("N_File_3")) <> "" then
	if UpLoad("N_File_3").FileLen > (1024 * 1024 * 10) then '10�ް� �������� üũ
		strError = "����3�� 10�ް����� ���ε� �����մϴ�.\n"
	end if
end if

rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

	N_File_1	= Trim(UpLoad("N_File_1"))
	oldN_File_1	= DefaultPath_Notice & Trim(UpLoad("oldN_File_1"))
	N_File_2	= Trim(UpLoad("N_File_2"))
	oldN_File_2	= DefaultPath_Notice & Trim(UpLoad("oldN_File_2"))
	N_File_3	= Trim(UpLoad("N_File_3"))
	oldN_File_3	= DefaultPath_Notice & Trim(UpLoad("oldN_File_3"))
	
	If N_File_1 <> "" then
		If oldN_File_1 <> "" Then
			File_Delete(oldN_File_1)
		End If
		N_File_1 = UpLoad("N_File_1").Save(,False)

	Else 
		If oldN_File_1 <> "" Then
			If InStr(strDelete, "N_File_1") > 0 Then
				File_Delete(oldN_File_1)
				N_File_1 = ""
			Else 
				N_File_1 = oldN_File_1
			End If 
		Else 
			N_File_1 = ""
		End If
	End If 

	If N_File_2 <> "" then
		If oldN_File_2 <> "" Then
			File_Delete(oldN_File_2)
		End If
		N_File_2 = UpLoad("N_File_2").Save(,False)
	Else 
		If oldN_File_2 <> "" Then
			If InStr(strDelete, "N_File_2") > 0 Then
				File_Delete(oldN_File_2)
				N_File_2 = ""
			Else 
				N_File_2 = oldN_File_2
			End If 
		Else 
			N_File_2 = ""
		End If
	End If 

	If N_File_3 <> "" then
		If oldN_File_3 <> "" Then
			File_Delete(oldN_File_3)
		End If
		N_File_3 = UpLoad("N_File_3").Save(,False)
	Else 
		If oldN_File_3 <> "" Then
			If InStr(strDelete, "N_File_3") > 0 Then
				File_Delete(oldN_File_3)
				N_File_3 = ""
			Else 
				N_File_3 = oldN_File_3
			End If 
		Else 
			N_File_3 = ""
		End If
	End If 

	N_Code		= UpLoad("N_Code")
	N_Title		= UpLoad("N_Title")
	N_Content	= UpLoad("N_Content")
	N_Edit_Date	= now()
	N_File_1 	= Replace(lcase(N_File_1),DefaultPath_Notice,"")
	N_File_2 	= Replace(lcase(N_File_2),DefaultPath_Notice,"")
	N_File_3 	= Replace(lcase(N_File_3),DefaultPath_Notice,"")
	
	rem DB ������Ʈ
	SQL = "select * from tbNotice where N_Code = '"&N_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1		
		.Fields("N_Title")		= N_Title
		.Fields("N_Content")	= N_Content
		.Fields("N_Edit_Date")	= N_Edit_Date
		.Fields("N_File_1")		= N_File_1
		.Fields("N_File_2")		= N_File_2
		.Fields("N_File_3")		= N_File_3
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
<input type="hidden" name="N_Code" value="<%=N_Code%>">

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
<input type="hidden" name="B_Code" value="<%=B_Code%>">

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