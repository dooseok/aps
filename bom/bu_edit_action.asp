<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<% 
rem ��������
dim SQL
dim RS1
dim UpLoad

dim BU_Code
dim BOM_B_D_No
dim BU_Content
dim BU_Receive_Date
dim BU_Apply_Date
dim BU_Reply_Date
dim BU_Request_Reply_Date
dim BU_File_1
dim BU_File_2
dim BU_File_3
dim BU_Type_New
dim BU_Type_Add
dim BU_Type_Update
dim BU_Type
dim Member_M_ID

dim oldBU_File_1
dim oldBU_File_2
dim oldBU_File_3

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
UpLoad.DefaultPath = DefaultPath_BOM_Update

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

strDelete	= UpLoad("strDelete")

BU_Code			= UpLoad("BU_Code")

SQL = "select Member_M_ID from tbBOM_Update where BU_Code = '"&UpLoad("BU_Code")&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	strError = strError & "*�ۼ��������� ȸ��DB���� ã�� �� �����ϴ�.\n*�����ڿ��� �����Ͽ� �ֽʽÿ�.\n"
else
	if lcase(RS1("Member_M_ID")) <> lcase(gM_ID) then
		'strError = strError & "*�ۼ��� ������ ������ ������ �� �ֽ��ϴ�.\n"
	end if
end if
RS1.Close

rem ���ε� �� ���� üũ
if trim(UpLoad("BU_File_1")) <> "" then
	if UpLoad("BU_File_1").FileLen > (1024 * 1024 * 50) then '10�ް� �������� üũ
		strError = "����1�� 50�ް����� ���ε� �����մϴ�.\n"
	end if
end if
if trim(UpLoad("BU_File_2")) <> "" then
	if UpLoad("BU_File_2").FileLen > (1024 * 1024 * 50) then '10�ް� �������� üũ
		strError = "����2�� 50�ް����� ���ε� �����մϴ�.\n"
	end if
end if
if trim(UpLoad("BU_File_3")) <> "" then
	if UpLoad("BU_File_3").FileLen > (1024 * 1024 * 50) then '10�ް� �������� üũ
		strError = "����3�� 50�ް����� ���ε� �����մϴ�.\n"
	end if
end if

rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

	BU_File_1	= Trim(UpLoad("BU_File_1"))
	oldBU_File_1	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_1"))
	BU_File_2	= Trim(UpLoad("BU_File_2"))
	oldBU_File_2	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_2"))
	BU_File_3	= Trim(UpLoad("BU_File_3"))
	oldBU_File_3	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_3"))
	
	If BU_File_1 <> "" then
		If oldBU_File_1 <> "" Then
			File_Delete(oldBU_File_1)
		End If
		BU_File_1 = UpLoad("BU_File_1").Save(,False)

	Else 
		If oldBU_File_1 <> "" Then
			If InStr(strDelete, "BU_File_1") > 0 Then
				File_Delete(oldBU_File_1)
				BU_File_1 = ""
			Else 
				BU_File_1 = oldBU_File_1
			End If 
		Else 
			BU_File_1 = ""
		End If
	End If 

	If BU_File_2 <> "" then
		If oldBU_File_2 <> "" Then
			File_Delete(oldBU_File_2)
		End If
		BU_File_2 = UpLoad("BU_File_2").Save(,False)
	Else 
		If oldBU_File_2 <> "" Then
			If InStr(strDelete, "BU_File_2") > 0 Then
				File_Delete(oldBU_File_2)
				BU_File_2 = ""
			Else 
				BU_File_2 = oldBU_File_2
			End If 
		Else 
			BU_File_2 = ""
		End If
	End If 

	If BU_File_3 <> "" then
		If oldBU_File_3 <> "" Then
			File_Delete(oldBU_File_3)
		End If
		BU_File_3 = UpLoad("BU_File_3").Save(,False)
	Else 
		If oldBU_File_3 <> "" Then
			If InStr(strDelete, "BU_File_3") > 0 Then
				File_Delete(oldBU_File_3)
				BU_File_3 = ""
			Else 
				BU_File_3 = oldBU_File_3
			End If 
		Else 
			BU_File_3 = ""
		End If
	End If 

	BU_Code			= UpLoad("BU_Code")
	BOM_B_D_No		= UpLoad("BOM_B_D_No")
	BU_Content		= UpLoad("BU_Content")
	BU_Receive_Date	= Trim(UpLoad("BU_Receive_Date"))
	BU_Apply_Date	= Trim(UpLoad("BU_Apply_Date"))
	BU_Reply_Date	= Trim(UpLoad("BU_Reply_Date"))
	BU_Request_Reply_Date	= Trim(UpLoad("BU_Request_Reply_Date"))
	BU_File_1 		= Replace(lcase(BU_File_1),DefaultPath_BOM_Update,"")
	BU_File_2 		= Replace(lcase(BU_File_2),DefaultPath_BOM_Update,"")
	BU_File_3 		= Replace(lcase(BU_File_3),DefaultPath_BOM_Update,"")
	BU_Type_New	= Trim(UpLoad("BU_Type_New"))
	BU_Type_Add	= Trim(UpLoad("BU_Type_Add"))
	BU_Type_Update	= Trim(UpLoad("BU_Type_Update"))
	
	if BU_Type_New = "Y" then
		BU_Type = BU_Type & "�ű�-"
	end if
	if BU_Type_Add = "Y" then
		BU_Type = BU_Type & "�߰�-"
	end if
	if BU_Type_Update = "Y" then
		BU_Type = BU_Type & "�ù�-"
	end if
	
	rem DB ������Ʈ
	SQL = "select * from tbBOM_Update where BU_Code = '"&BU_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1		
		.Fields("BU_Content")	= BU_Content
		If BU_Receive_Date <> "" then
			.Fields("BU_Receive_Date")	= BU_Receive_Date
		End If
		If BU_Apply_Date <> "" then
			.Fields("BU_Apply_Date")	= BU_Apply_Date
		End If
		If BU_Reply_Date <> "" then
			.Fields("BU_Reply_Date")	= BU_Reply_Date
		End If
		If BU_Request_Reply_Date <> "" then
			.Fields("BU_Request_Reply_Date")	= BU_Request_Reply_Date
		End if
		.Fields("BU_File_1")	= BU_File_1
		.Fields("BU_File_2")	= BU_File_2
		.Fields("BU_File_3")	= BU_File_3
		.Fields("BOM_B_D_No")	= BOM_B_D_No
		.Fields("BU_Type")		= BU_Type
		.Update
		.Close
	end with
	
	Member_M_ID		= gM_ID
	rem DB ������Ʈ
	RS1.Open "tbNotice",sys_DBConString,3,2,2
	with RS1
		.AddNew
		if BOM_B_D_No <> "" then
			.Fields("N_Title")			= BOM_B_D_No & "�� ���� �ù��� ���� ��ϵǾ����ϴ�."
			.Fields("N_Content")		= BU_Content
		else
			.Fields("N_Title")			= "�ù��� ��ϵǾ����ϴ�."
			.Fields("N_Content")		= BU_Content
		end if
		
		.Fields("N_Reg_Date")		= date()
		.Fields("N_Edit_Date")		= date()
		.Fields("N_File_1")			= ""
		.Fields("N_File_2")			= ""
		.Fields("N_File_3")			= ""
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
<form name="frmRedirect" action="bu_edit_form.asp" method=post>
<input type="hidden" name="BU_Code" value="<%=BU_Code%>">
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
<form name="frmRedirect" action="bu_edit_form.asp" method=post>
<input type="hidden" name="BU_Code" value="<%=BU_Code%>">
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