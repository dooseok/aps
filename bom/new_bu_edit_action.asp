<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->

<!-- #include Virtual = "/bom/new_bu_file_partno_update.asp" -->

<form name='frmErrorRedirect' action='new_bu_list.asp' method=post></form>

<% 
rem ��������
dim SQL
dim RS1
dim UpLoad


dim BU_LG_Part
dim BU_LG_Staff
dim BU_Eco_No
dim BU_Sibang_No
dim BU_Parts_PNO
dim BU_Last_Use_Date

dim BU_Code
dim BOM_B_D_No
dim BU_Content
dim BU_Receive_Date
dim BU_Apply_Date

dim BU_MSE_LG
dim BU_Link_YN

dim BU_File_PartNo
dim BU_File_1
dim BU_File_2
dim BU_File_3
dim BU_File_4
dim BU_File_5
dim BU_Type_SW
dim BU_Type_HW
dim BU_Type_REAL
dim BU_Type_SAMPLE
dim BU_Type_New
dim BU_Type_Add
dim BU_Type_Update
dim BU_Type

dim Member_M_ID

dim BU_RnD_Check
dim BU_RnD_Memo

dim BU_RnD_Check_Old

dim oldBU_File_PartNo
dim oldBU_File_1
dim oldBU_File_2
dim oldBU_File_3
dim oldBU_File_4
dim oldBU_File_5
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
UpLoad.MaxFileLen = (1024 * 1024 * 50)

URL_Prev	= UpLoad("URL_Prev")
URL_Next	= UpLoad("URL_Next")

strDelete	= UpLoad("strDelete")

BU_Code		= UpLoad("BU_Code")

SQL = "select Member_M_ID from tbBOM_Update_New where BU_Code = '"&UpLoad("BU_Code")&"'"
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
if trim(UpLoad("BU_File_PartNo")) <> "" then
	if UpLoad("BU_File_PartNo").FileLen > Upload.MaxFileLen then '200�ް� �������� üũ
		strError = "����(ǰ��)�� 50�ް����� ���ε� �����մϴ�.\n"
%><script>alert("<%=strError%>");frmRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("BU_File_1")) <> "" then
	if UpLoad("BU_File_1").FileLen > Upload.MaxFileLen then '200�ް� �������� üũ
		strError = "����1�� 50�ް����� ���ε� �����մϴ�.\n"
%><script>alert("<%=strError%>");frmRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("BU_File_2")) <> "" then
	if UpLoad("BU_File_2").FileLen > Upload.MaxFileLen then '200�ް� �������� üũ
		strError = "����2�� 50�ް����� ���ε� �����մϴ�.\n"
%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("BU_File_3")) <> "" then
	if UpLoad("BU_File_3").FileLen > Upload.MaxFileLen then '200�ް� �������� üũ
		strError = "����3�� 50�ް����� ���ε� �����մϴ�.\n"
%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("BU_File_4")) <> "" then
	if UpLoad("BU_File_4").FileLen > Upload.MaxFileLen then '200�ް� �������� üũ
		strError = "����4�� 50�ް����� ���ε� �����մϴ�.\n"
%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
if trim(UpLoad("BU_File_5")) <> "" then
	if UpLoad("BU_File_5").FileLen > Upload.MaxFileLen then '200�ް� �������� üũ
		strError = "����5�� 50�ް����� ���ε� �����մϴ�.\n"
%><script>alert("<%=strError%>");frmErrorRedirect.submit();</script><%
		response.end
	end if
end if
rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

	BU_File_PartNo	= Trim(UpLoad("BU_File_PartNo"))
	oldBU_File_PartNo	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_PartNo"))
	BU_File_1	= Trim(UpLoad("BU_File_1"))
	oldBU_File_1	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_1"))
	BU_File_2	= Trim(UpLoad("BU_File_2"))
	oldBU_File_2	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_2"))
	BU_File_3	= Trim(UpLoad("BU_File_3"))
	oldBU_File_3	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_3"))
	BU_File_4	= Trim(UpLoad("BU_File_4"))
	oldBU_File_4	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_4"))
	BU_File_5	= Trim(UpLoad("BU_File_5"))
	oldBU_File_5	= DefaultPath_BOM_Update & Trim(UpLoad("oldBU_File_5"))
	
	If BU_File_PartNo <> "" then
		If oldBU_File_PartNo <> "" Then
			File_Delete(oldBU_File_PartNo)
			call BU_File_PartNo_Update(BU_Code, "")
		End If
		BU_File_PartNo = UpLoad("BU_File_PartNo").Save(,False)
		call BU_File_PartNo_Update(BU_Code, BU_File_PartNo)
	Else 
		If oldBU_File_PartNo <> "" Then
			If InStr(strDelete, "BU_File_PartNo") > 0 Then
				File_Delete(oldBU_File_PartNo)
				call BU_File_PartNo_Update(BU_Code, "")
				BU_File_PartNo = ""
			Else 
				BU_File_PartNo = oldBU_File_PartNo
			End If 
		Else 
			BU_File_PartNo = ""
		End If
	End If 
	
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
	
	If BU_File_4 <> "" then
		If oldBU_File_4 <> "" Then
			File_Delete(oldBU_File_4)
		End If
		BU_File_4 = UpLoad("BU_File_4").Save(,False)
	Else 
		If oldBU_File_4 <> "" Then
			If InStr(strDelete, "BU_File_4") > 0 Then
				File_Delete(oldBU_File_4)
				BU_File_4 = ""
			Else 
				BU_File_4 = oldBU_File_4
			End If 
		Else 
			BU_File_4 = ""
		End If
	End If 
	
	If BU_File_5 <> "" then
		If oldBU_File_5 <> "" Then
			File_Delete(oldBU_File_5)
		End If
		BU_File_5 = UpLoad("BU_File_5").Save(,False)
	Else 
		If oldBU_File_5 <> "" Then
			If InStr(strDelete, "BU_File_5") > 0 Then
				File_Delete(oldBU_File_5)
				BU_File_5 = ""
			Else 
				BU_File_5 = oldBU_File_5
			End If 
		Else 
			BU_File_5 = ""
		End If
	End If 
	
	BU_LG_Part		= Trim(UpLoad("BU_LG_Part"))
	BU_LG_Staff		= Trim(UpLoad("BU_LG_Staff"))
	BU_Eco_No		= Trim(UpLoad("BU_Eco_No"))
	BU_Sibang_No	= Trim(UpLoad("BU_Sibang_No"))
	BU_Parts_PNO	= Trim(UpLoad("BU_Parts_PNO"))
	BU_Last_Use_Date= Trim(UpLoad("BU_Last_Use_Date"))
	
	BU_Code			= UpLoad("BU_Code")
	BOM_B_D_No		= UpLoad("BOM_B_D_No")
	BU_Content		= UpLoad("BU_Content")
	BU_Receive_Date	= Trim(UpLoad("BU_Receive_Date"))
	BU_Apply_Date	= Trim(UpLoad("BU_Apply_Date"))
	
	BU_MSE_LG	= Trim(UpLoad("BU_MSE_LG"))
	BU_Link_YN	= Trim(UpLoad("BU_Link_YN"))
	
	BU_File_PartNo	= Replace(lcase(BU_File_PartNo),DefaultPath_BOM_Update,"")
	BU_File_1 		= Replace(lcase(BU_File_1),DefaultPath_BOM_Update,"")
	BU_File_2 		= Replace(lcase(BU_File_2),DefaultPath_BOM_Update,"")
	BU_File_3 		= Replace(lcase(BU_File_3),DefaultPath_BOM_Update,"")
	BU_File_4 		= Replace(lcase(BU_File_4),DefaultPath_BOM_Update,"")
	BU_File_5 		= Replace(lcase(BU_File_5),DefaultPath_BOM_Update,"")
	BU_Type_SW	= Trim(UpLoad("BU_Type_SW"))
	BU_Type_HW	= Trim(UpLoad("BU_Type_HW"))
	BU_Type_REAL	= Trim(UpLoad("BU_Type_REAL"))
	BU_Type_SAMPLE	= Trim(UpLoad("BU_Type_SAMPLE"))
	BU_Type_New	= Trim(UpLoad("BU_Type_New"))
	BU_Type_Add	= Trim(UpLoad("BU_Type_Add"))
	BU_Type_Update	= Trim(UpLoad("BU_Type_Update"))
	
	
	if BU_Type_SW = "Y" then
		BU_Type = BU_Type & "S/W-"
	end if
	if BU_Type_HW = "Y" then
		BU_Type = BU_Type & "H/W-"
	end if
	if BU_Type_REAL = "Y" then
		BU_Type = BU_Type & "����ȭ-"
	end if
	if BU_Type_SAMPLE = "Y" then
		BU_Type = BU_Type & "�������-"
	end if
	if BU_Type_New = "Y" then
		BU_Type = BU_Type & "�ű�-"
	end if
	if BU_Type_Add = "Y" then
		BU_Type = BU_Type & "�߰�-"
	end if
	if BU_Type_Update = "Y" then
		BU_Type = BU_Type & "�ù�-"
	end if
	
	
	BU_RnD_Check	= trim(UpLoad("BU_RnD_Check"))
	BU_Rnd_Memo		= trim(UpLoad("BU_Rnd_Memo"))
	
	
	
	'BU_Code�� �ٽ� ����.
	dim oldBU_RnD_Check
	SQL = "select BU_RnD_Check from tbBOM_Update_New where BU_Code = '"&BU_Code&"'"
	RS1.Open SQL,sys_DBCon
	oldBU_RnD_Check = RS1("BU_RnD_Check")
	RS1.Close
	
	
	'Ȯ��or��Ȯ�� <--> �ش����or����Ϸ� �� �ٲ� ��, BU_Code�� ���ξ ��ȯ���ش�. 
	dim newBU_Code
	dim cntBU_Code
	if (instr("�ش����,����Ϸ�",BU_RnD_Check) > 0 and instr("Ȯ��,��Ȯ��",oldBU_RnD_Check) > 0) or (instr("Ȯ��,��Ȯ��",BU_RnD_Check) > 0 and instr("�ش����,����Ϸ�",oldBU_RnD_Check) > 0 ) then
		SQL = "select max(convert(integer,right(bu_code,3))) from tbBOM_Update_New where '20'+substring(BU_Code,3,5)+'-'+substring(BU_Code,8,2) = '"&date()&"' and BU_Code <> '"&BU_Code&"'"
		RS1.Open SQL,sys_DBCon
		if RS1.Eof or RS1.Bof then
			cntBU_Code = 0
		elseif isnull(RS1(0)) then
			cntBU_Code = 0
		elseif not(isnumeric(RS1(0))) then
			cntBU_code = 0
		else	
			cntBU_Code = RS1(0)
		end if
		RS1.Close
		cntBU_Code = cntBU_Code + 1
		if len(cntBU_Code)=1 then
			cntBU_Code = "00" & cntBU_Code
		elseif len(cntBU_Code)=2 then
			cntBU_Code = "0" & cntBU_Code
		end if
		newBU_Code = mid(date(),3,5) & right(date(),2) & "-" & cntBU_Code
		if instr("�ش����,����Ϸ�",BU_RnD_Check) > 0 and instr("Ȯ��,��Ȯ��",oldBU_RnD_Check) > 0 then
			newBU_Code = "MS"&newBU_Code
		elseif instr("Ȯ��,��Ȯ��",BU_RnD_Check) > 0 and instr("�ش����,����Ϸ�",oldBU_RnD_Check) > 0 then
			SQL = "update tbBOM_Update_New set "
			SQL = SQL & "BU_JaJe_Check	 = '��Ȯ��', BU_JaJe_Date = null, "
			SQL = SQL & "BU_IMT_Check	 = '��Ȯ��', BU_IMT_Date = null, "
			SQL = SQL & "BU_SMT_Check	 = '��Ȯ��', BU_SMT_Date = null, "
			SQL = SQL & "BU_JeJo2_Check	 = '��Ȯ��', BU_JeJo2_Date = null, "
			SQL = SQL & "BU_JeJo3_Check	 = '��Ȯ��', BU_JeJo3_Date = null, "
			SQL = SQL & "BU_IQC_Check	 = '��Ȯ��', BU_IQC_Date = null, "
			SQL = SQL & "BU_PCBA_QC_Check= '��Ȯ��', BU_PCBA_QC_Date = null, "
			SQL = SQL & "BU_CBOX_QC_Check= '��Ȯ��', BU_CBOX_QC_Date = null, "
			SQL = SQL & "BU_SPMK_Check	 = '��Ȯ��', BU_SPMK_Date = null, "
			SQL = SQL & "BU_DLV_Check	 = '��Ȯ��', BU_DLV_Date = null, "
			SQL = SQL & "BU_Price_Check	 = '��Ȯ��', BU_Price_Date = null, "
			SQL = SQL & "BU_OTP_Check	 = '��Ȯ��', BU_OTP_Date = null, "
			SQL = SQL & "BU_Eng_Check	 = '��Ȯ��', BU_Eng_Date = null, "
			SQL = SQL & "BU_SMTech_Check = '��Ȯ��', BU_SMTech_Date = null, "
			SQL = SQL & "BU_DSTech_Check = '��Ȯ��', BU_DSTech_Date = null "
			SQL = SQL & "where BU_Code = '"&BU_Code&"'"
			sys_DBCon.execute(SQL)
			
			newBU_Code = "RS"&newBU_Code
		end if
	
	 	SQL = "update tbBOM_Update_New set BU_Code = '"&newBU_Code&"' where BU_Code = '"&BU_Code&"'"
		sys_DBCon.execute(SQL)
		
		SQL = "update tbBOM_Update_PartNo set BOM_Update_BU_Code = '"&newBU_Code&"' where BOM_Update_BU_Code = '"&BU_Code&"'"
		sys_DBCon.execute(SQL)
		
		BU_Code = newBU_Code
	end if
	
	SQL = "select BU_RnD_Check from tbBOM_Update_New where BU_Code = '"&BU_Code&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		BU_RnD_Check_Old = RS1("BU_RnD_Check")
	end if
	RS1.Close
	
	rem DB ������Ʈ
	SQL = "select * from tbBOM_Update_New where BU_Code = '"&BU_Code&"'"
	RS1.Open SQL,sys_DBconString,3,2,&H0001
	with RS1		
	
		.Fields("BU_LG_Part")		= BU_LG_Part
		.Fields("BU_LG_Staff")		= BU_LG_Staff
		.Fields("BU_Eco_No")		= BU_Eco_No
		.Fields("BU_Sibang_No")		= BU_Sibang_No
		.Fields("BU_Parts_PNO")		= BU_Parts_PNO
		.Fields("BU_Last_Use_Date")	= BU_Last_Use_Date
	
		.Fields("BU_Content")	= BU_Content
		If BU_Receive_Date <> "" then
			.Fields("BU_Receive_Date")	= BU_Receive_Date
		End If
		If BU_Apply_Date <> "" then
			.Fields("BU_Apply_Date")	= BU_Apply_Date
		End If
		.Fields("BU_MSE_LG")	= BU_MSE_LG
		.Fields("BU_Link_YN")	= BU_Link_YN
		
		.Fields("BU_File_PartNo")= BU_File_PartNo
		.Fields("BU_File_1")	= BU_File_1
		.Fields("BU_File_2")	= BU_File_2
		.Fields("BU_File_3")	= BU_File_3
		.Fields("BU_File_4")	= BU_File_4
		.Fields("BU_File_5")	= BU_File_5
		.Fields("BOM_B_D_No")	= BOM_B_D_No
		.Fields("BU_Type")		= BU_Type
		
		.Fields("BU_RnD_Check")	= BU_RnD_Check
		.Fields("BU_Rnd_Memo")	= BU_Rnd_Memo
		if BU_RnD_Check="��Ȯ��" then
			.Fields("BU_RnD_Date") = null 
		else
			.Fields("BU_RnD_Date") = date()
		end if
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
	
	'�� ������ �������� �Խ��ǿ��� ��üǰ���� ���� �Ѵ�. -230331 ����
	'���°� ���� ���� !�ߺ����� ��� ���� �ʵ��� �Ѵ�!!
	if BOM_B_D_No <> "" and BU_RnD_Check = "����Ϸ�" and BU_RnD_Check_Old <> "����Ϸ�" then
		SQL = "update tblBOM_Process_Info set "
		SQL = SQL & "BPI_PNO_EYELET = '', "
		SQL = SQL & "BPI_PNO_SMD_T = '', "
		SQL = SQL & "BPI_PNO_SMD_B = '', "
		SQL = SQL & "BPI_PNO_AXIAL = '', "
		SQL = SQL & "BPI_PNO_RADIAL = '', "
		SQL = SQL & "BPI_PNO_PCBA = '', "
		SQL = SQL & "BPI_STATE = 'CHECK' "
		SQL = SQL & "where "
		SQL = SQL & "(BPI_PNO_EYELET <> '' or "
		SQL = SQL & "BPI_PNO_SMD_T <> '' or "
		SQL = SQL & "BPI_PNO_SMD_B <> '' or "
		SQL = SQL & "BPI_PNO_AXIAL <> '' or "
		SQL = SQL & "BPI_PNO_RADIAL <> '' or "
		SQL = SQL & "BPI_PNO_PCBA <> '') and "
		SQL = SQL & "BOM_Sub_BS_D_No in (select BS_D_No from tbBOM_Sub where BOM_B_Code = (select top 1 B_Code from tbBOM where B_D_No = '"&BOM_B_D_No&"' and B_Version_Current_YN='Y'))"
		sys_DBCon.execute(SQL)
	end if
	
	'reg_action�� �޸�, edit_action�� BOM_Update ���̺� �������� ����ó���� �ϹǷ�, �ϱ�� ���� ���� ������Ʈ �۾��� ��
	SQL = "update tbBOM_Update_PartNo set "
	if BU_Apply_Date <> "" then
		SQL = SQL & "	BOM_Update_BU_Apply_Date = '"&BU_Apply_Date&"', "
	end if
	SQL = SQL & "	BOM_Update_BU_MSE_LG = '"&BU_MSE_LG&"', "
	SQL = SQL & "	BOM_Update_BU_Link_YN = '"&BU_Link_YN&"' "
	SQL = SQL & "where BOM_Update_BU_Code = '"&BU_Code&"'"
	sys_DBCon.execute(SQL)
end if

rem ��ü ����
set UpLoad	= nothing
Set RS1		= nothing
%>

<form name="frmRedirect" action="new_bu_edit_form.asp" method=post>
<input type="hidden" name="BU_Code" value="<%=BU_Code%>">
<%
response.write strRequestForm
%>
</form>

<script language="javascript">
frmRedirect.submit();
</script>


<!-- #include Virtual = "/header/db_tail.asp" -->