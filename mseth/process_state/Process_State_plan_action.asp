<!-- #include Virtual = "/mseth/header/asp_header.asp" -->
<!-- include Virtual = "/mseth/header/session_check_header.asp" -->
<!-- #include Virtual = "/mseth/header/db_header.asp" -->
<!-- #include Virtual = "/mseth/header/inc_share_function.asp" -->
<%
'���� ����
dim CNT1
dim CNT2
dim RS1
dim SQL

'������ �ʿ� �÷���
dim ReSave_Require_YN

'SQL���ڿ��� 
dim strBS_ST_Update_PNO
dim strBS_ST_Update_SQL
dim arrBS_ST_Update_SQL

'���� ���ڿ�
dim strError

'��ȹ ��¥, ���� ����
dim s_Work_Date
dim s_Line

'����Ʈ�� ���ڿ� �迭
dim strBOM_Sub_BS_D_No
dim strPSP_Count
dim strPSP_ST
dim strPSP_Desc
dim strPSP_Start
dim strPSP_End
dim strPSP_Sub_Start
dim strPSP_Sub_End

'����Ʈ�� �迭
dim arrBOM_Sub_BS_D_No
dim arrPSP_Count
dim arrPSP_ST
dim arrPSP_Desc
dim arrPSP_Start
dim arrPSP_End
dim arrPSP_Sub_Start
dim arrPSP_Sub_End

dim PSP_Sub_YN
dim strPSP_Sub_YN
dim arrPSP_Sub_YN

'textarea�� ���� �޴� ���
dim lstBOM_Sub_BS_D_No
dim lstPSP_Count
dim lstPSP_ST

set RS1 = Server.CreateObject("ADODB.RecordSet")

strPSP_Sub_YN = ", "&request("PSP_Sub_YN")&","

lstBOM_Sub_BS_D_No	= trim(request("lstBOM_Sub_BS_D_No"))
lstPSP_Count		= trim(request("lstPSP_Count"))
lstPSP_ST			= trim(request("lstPSP_ST"))

'textarea �������� �޾Ƽ�, ����Ʈ ������ȭ ��Ŵ
strBOM_Sub_BS_D_No	= replace(lstBOM_Sub_BS_D_No	,chr(13)&chr(10),",")
strPSP_Count		= replace(lstPSP_Count			,chr(13)&chr(10),",")
strPSP_ST			= replace(lstPSP_ST				,chr(13)&chr(10),",")

s_Work_Date			= request("s_Work_Date")
s_Line				= request("s_Line")

'textarea �������� ���ٸ� > �� ����Ʈ �������̶��
if strBOM_Sub_BS_D_No = "" then
	strBOM_Sub_BS_D_No	= request("BOM_Sub_BS_D_No")
	strPSP_Count		= request("PSP_Count")
	strPSP_ST			= request("PSP_ST")
	strPSP_Desc			= request("PSP_Desc")
	strPSP_Start		= request("PSP_Start")
	strPSP_End			= request("PSP_End")
	strPSP_Sub_Start	= request("PSP_Sub_Start")
	strPSP_Sub_End		= request("PSP_Sub_End")
end if

'�켱 ��Ʈ�ѹ�, ����, st�� �迭ȭ
arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,",")
arrPSP_Count		= split(strPSP_Count,",")
arrPSP_ST			= split(strPSP_ST,",")

'textarea�� st�� ���Ծ��ٸ�, �� �迭�� �����.
if trim(lstBOM_Sub_BS_D_No) <> "" and trim(lstPSP_ST) = "" then
	redim arrPSP_ST(ubound(arrBOM_Sub_BS_D_No))
end if

'textarea�� ���� �����, �� �迭 ����ְ�, ����Ʈ�� ���� ��� �迭ȭ
if lstBOM_Sub_BS_D_No <> "" then
	redim arrPSP_Desc(ubound(arrBOM_Sub_BS_D_No))
	redim arrPSP_Start(ubound(arrBOM_Sub_BS_D_No))
	redim arrPSP_End(ubound(arrBOM_Sub_BS_D_No))
	redim arrPSP_Sub_Start(ubound(arrBOM_Sub_BS_D_No))
	redim arrPSP_Sub_End(ubound(arrBOM_Sub_BS_D_No))
else
	arrPSP_Desc			= split(strPSP_Desc,",")
	arrPSP_Start		= split(strPSP_Start,",")
	arrPSP_End			= split(strPSP_End,",")
	arrPSP_Sub_Start	= split(strPSP_Sub_Start,",")
	arrPSP_Sub_End		= split(strPSP_Sub_End,",")
end if

'���� ��ȹ ������ �����Ѵ�.
SQL = "delete tbProcess_State_Plan where PSP_Work_Date = '"&s_Work_Date&"' and PSP_Line = '"&s_Line&"'"
sys_DBCon.execute(SQL)

'st�� db�� �ٸ� ��찡 �ִٸ� ������Ʈ�� �ؾ� �ϱ� ������, �α� ���ڿ��� ����Ѵ�. �켱 �ʱ�ȭ
strBS_ST_Update_SQL = ""
strBS_ST_Update_PNO = "-"

'�迭�� ����
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
  '��Ʈ�ѹ��� ��ȿ�ϴٸ�
	if trim(arrBOM_Sub_BS_D_No(CNT1)) <> "" and len(trim(arrBOM_Sub_BS_D_No(CNT1)))=11 and isnumeric(arrPSP_Count(CNT1)) then
		'���� ���� �� ��Ʈ�ѹ� �빮�� ��ȯ
		arrBOM_Sub_BS_D_No(CNT1)	= ucase(trim(arrBOM_Sub_BS_D_No(CNT1)))
		arrPSP_ST(CNT1)						= trim(arrPSP_ST(CNT1))
		
				
		'db���� �ش� ��Ʈ�ѹ��� st���� ã�´�.
		SQL = "select BS_ST, BS_ST_ASM from tbBOM_Sub where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'"
		RS1.Open SQL,sys_DBCon
		
		'��ġ�ϴ� ��Ʈ�ѹ��� ���ٸ�.
		if RS1.Eof or RS1.Bof then
			'ST���� �� �Դٸ�, 10�ʷ� ��
			if arrPSP_ST(CNT1) = "" or not(isnumeric(arrPSP_ST(CNT1))) then
				arrPSP_ST(CNT1) = "10"
			end if
			'�� ��Ʈ�ѹ��� ������ st������ ����Ѵ�.
			SQL = "insert into tbBOM_Sub (BS_D_No, BS_ST, BS_ST_ASM) values ('"&arrBOM_Sub_BS_D_No(CNT1)&"',"&arrPSP_ST(CNT1)&","&arrPSP_ST(CNT1)&")"
			sys_DBCon.execute(SQL)

		'��ġ�ϴ� ��Ʈ�ѹ��� �ִٸ�                                                                                                                                                                                                       
		else
			'���� �Էµ� ST�� �����ϸ� DB���� ����, �������ϸ�, DB�� ST�� Ȱ��
			if not(isnumeric(arrPSP_ST(CNT1))) then
				arrPSP_ST(CNT1) = RS1("BS_ST")
			end if
			
			if instr(strBS_ST_Update_PNO,"-"&arrBOM_Sub_BS_D_No(CNT1)&"-") = 0 then '�ѹ� �̻� SQL���� ���Ե� PNO��� ����'
				if isnull(RS1("BS_ST")) or isnull(RS1("BS_ST_ASM")) then 'DB���� ST������ null�̸�. ���ο� ST�� ������Ʈ
					strBS_ST_Update_SQL = strBS_ST_Update_SQL & "update tbBOM_Sub set BS_ST = "&arrPSP_ST(CNT1)&", BS_ST_ASM = "&arrPSP_ST(CNT1)&" where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'-----"
					strBS_ST_Update_PNO = strBS_ST_Update_PNO & arrBOM_Sub_BS_D_No(CNT1) & "-"
				elseif int(RS1("BS_ST")) <> int(arrPSP_ST(CNT1)) or int(RS1("BS_ST_ASM")) <> int(arrPSP_ST(CNT1)) then 'DB���� ������ �����ϴٸ�, ���ο� ST�� ������Ʈ'
					strBS_ST_Update_SQL = strBS_ST_Update_SQL & "update tbBOM_Sub set BS_ST = "&arrPSP_ST(CNT1)&", BS_ST_ASM = "&arrPSP_ST(CNT1)&" where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'-----"
					strBS_ST_Update_PNO = strBS_ST_Update_PNO & arrBOM_Sub_BS_D_No(CNT1) & "-"
				end if
			end if
		end if
		RS1.Close
	end if
next

'���Ƽ� ������Ʈ
arrBS_ST_Update_SQL = split(strBS_ST_Update_SQL,"-----")
for CNT1=0 to ubound(arrBS_ST_Update_SQL)-1
	sys_DBCon.execute(arrBS_ST_Update_SQL(CNT1))
next

'�켱 ������ �÷��׸� N���� �д�.
ReSave_Require_YN = "N"

'�ٽ� ����
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
	'��Ʈ�ѹ��� ��ȿ�ϸ�,
	if trim(arrBOM_Sub_BS_D_No(CNT1)) <> "" and len(trim(arrBOM_Sub_BS_D_No(CNT1)))=11 then
		'��������
		arrBOM_Sub_BS_D_No(CNT1)	= trim(arrBOM_Sub_BS_D_No(CNT1))
		arrPSP_Count(CNT1)			= trim(arrPSP_Count(CNT1))
		arrPSP_ST(CNT1)				= trim(arrPSP_ST(CNT1))
		arrPSP_Desc(CNT1)			= trim(arrPSP_Desc(CNT1))
		arrPSP_Start(CNT1)			= trim(arrPSP_Start(CNT1))
		arrPSP_End(CNT1)			= trim(arrPSP_End(CNT1))
		arrPSP_Sub_Start(CNT1)		= trim(arrPSP_Sub_Start(CNT1))
		arrPSP_Sub_End(CNT1)		= trim(arrPSP_Sub_End(CNT1))
		
		'ST�� ���ڰ� �ƴϸ� 10�ʷ� ��
		if not(isnumeric(arrPSP_ST(CNT1))) then
			arrPSP_ST(CNT1) = 10
		end if
		
		'��Ʈ�ѹ��� �ִµ�, ���۽ð��� ���ٸ�, ������ �÷��� Y
		if arrBOM_Sub_BS_D_No(CNT1) <> "" and trim(arrPSP_Start(CNT1)) = "" then
			ReSave_Require_YN = "Y"
		end if 
		
		'��Ʈ�ѹ� ��ȿ�ϰ�, ��ǥ������ ��ȿ�ϴٸ�
		if arrBOM_Sub_BS_D_No(CNT1) <> "" and isnumeric(arrPSP_Count(CNT1)) then
			'��ǥ������ ���� 0���� ũ�� ��ȿ�ϴٸ�,
			if arrPSP_Count(CNT1) > 0 then
				'DB�� ST�� ������
				SQL = "select BS_ST, BS_ST_ASM from tbBOM_Sub where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'"
				RS1.Open SQL,sys_DBCon
				if RS1.Eof or RS1.Bof then '��ġ�ϴ� ��Ʈ�ѹ� ������ st 10�ʷ� ����
					arrPSP_ST(CNT1) = 10
				else
					arrPSP_ST(CNT1) = RS1("BS_ST") '��ġ�ϴ� ��Ʈ�ѹ� ������ �� ���� ������.
				end if
				RS1.Close
				
				'����PCB����
				PSP_Sub_YN = ""
				if instr(strPSP_Sub_YN,", "&cstr(CNT1)&",") > 0 then
					PSP_Sub_YN = "Y"
				end if
				
				'��ȹDB�� ����
				SQL = "insert into tbProcess_State_Plan (BOM_Sub_BS_D_No, PSP_Count, PSP_ST, PSP_Desc, PSP_Start, PSP_End, PSP_Sub_YN, PSP_Sub_Start, PSP_Sub_End, PSP_Work_Date, PSP_Line) values "
				SQL = SQL & "('"&arrBOM_Sub_BS_D_No(CNT1)&"',"&arrPSP_Count(CNT1)&","&arrPSP_ST(CNT1)&",'"&arrPSP_Desc(CNT1)&"','"&arrPSP_Start(CNT1)&"','"&arrPSP_End(CNT1)&"','"&PSP_Sub_YN&"','"&arrPSP_Sub_Start(CNT1)&"','"&arrPSP_Sub_End(CNT1)&"','"&s_Work_Date&"','"&s_Line&"')"		
				sys_DBCon.execute(SQL)
			end if
		end if
	end if
next
set RS1 = nothing

SQL = "delete from tbPWS_Raw_Data where PRD_Input_Date < '"&dateadd("m",-2,date())&"'"
sys_DBCon.execute(SQL)

if strError = "" then
%>
<form name="frmRedirect" action="Process_State_Plan.asp" method=post>
<input type="hidden" name="s_Work_Date"			value="<%=request("s_Work_Date")%>">
<input type="hidden" name="s_Line"				value="<%=request("s_Line")%>">
<input type="hidden" name="s_ReSave_Require_YN"	value="<%=ReSave_Require_YN%>">
<input type="hidden" name="s_WorkStart"				value="<%=request("s_WorkStart")%>">
</form>
<script language="javascript">
//parent.ifrmRecord.location.reload();
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="Process_State_Plan.asp" method=post>
<input type="hidden" name="s_Work_Date"			value="<%=request("s_Work_Date")%>">
<input type="hidden" name="s_Line"				value="<%=request("s_Line")%>">
<input type="hidden" name="s_ReSave_Require_YN"	value="<%=ReSave_Require_YN%>">
<input type="hidden" name="s_WorkStart"				value="<%=request("s_WorkStart")%>">
</form>
<script language="javascript">
alert("<%=strError%>");
//parent.ifrmRecord.location.reload();
frmRedirect.submit();
</script>
<%
end if
%>



<!-- #include Virtual = "/mseth/header/db_tail.asp" -->
<!-- include Virtual = "/mseth/header/session_check_tail.asp" -->