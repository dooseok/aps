<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->


<%
'DB�� ���� ����
dim RS1
dim SQL

'�ݺ����� ����ϱ� ���� ���� ����
dim CNT1
dim CNT2
dim CNT3


dim strBOM_Sub_BS_D_No	'��Ʈ�ѹ�
dim strPSP_Count		'��ǥ����
dim strPSP_ST			'ǥ�ؽð�
dim strPSP_Desc			'���
dim strPSP_Start		'���۽ð�
dim strPSP_End			'����ð�

dim arrBOM_Sub_BS_D_No	
dim arrPSP_Count		
dim arrPSP_ST			
dim arrPSP_Desc			
dim arrPSP_Start		
dim arrPSP_End			


dim strPSP_Sub_YN		'����PCB �÷���
dim strPSP_Sub_Start	'���۽ð�
dim strPSP_Sub_End		'����ð�

dim arrPSP_Sub_YN		
dim arrPSP_Sub_Start	
dim arrPSP_Sub_End		

'����ȭ �� �迭
dim arrOptBOM_Sub_BS_D_No(60)
dim arrOptPSP_Count(60)
dim arrOptPSP_ST(60)
dim arrOptPSP_Desc(60)
dim arrOptPSP_Start(60)
dim arrOptPSP_End(60)

dim arrOptPSP_Period(60)

dim oldBOM_Sub_BS_D_No	
dim BOM_Sub_BS_D_No
dim PSP_Count
dim PSP_ST
dim PSP_Start
dim PSP_End
dim PSP_Desc
dim PSP_Sub_YN			
dim PSP_Sub_Start		
dim PSP_Sub_End	

'��ü���� �ð�
dim nMC_Time

'�ҿ�ð�
dim PSP_Period

'��ȿ �� ����
dim nDBRowLength

'������ �ð�ǥ��
dim nAccTime

'������ ���� �ð� ǥ��
dim nRest

'���� �ð� ����Ÿ �Ҵ�
dim arrRest(3,1)

dim PS	'��ȹ ����
dim PE	'��ȹ ����
dim RS	'�޽� ����
dim RE	'�޽� ����

'��ȹ�� DB���� ��ȸ, ���ΰ� ��¥�� ��ȸ
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbProcess_State_Plan where PSP_Line = '"&Request("s_Line")&"' and PSP_Work_Date = '"&Request("S_Work_Date")&"' order by PSP_Code asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strBOM_Sub_BS_D_No	= strBOM_Sub_BS_D_No 	& RS1("BOM_Sub_BS_D_No")	& ","
	strPSP_Count		= strPSP_Count 			& RS1("PSP_Count")			& ","
	strPSP_ST			= strPSP_ST 			& RS1("PSP_ST")				& ","
	strPSP_Desc			= strPSP_Desc			& RS1("PSP_Desc")			& ","
	strPSP_Start		= strPSP_Start 			& RS1("PSP_Start")			& ","
	strPSP_End			= strPSP_End			& RS1("PSP_End")			& ","
	strPSP_Sub_YN		= strPSP_Sub_YN			& RS1("PSP_Sub_YN")			& ","
	strPSP_Sub_Start	= strPSP_Sub_Start 		& RS1("PSP_Sub_Start")		& ","
	strPSP_Sub_End		= strPSP_Sub_End		& RS1("PSP_Sub_End")		& ","
	RS1.MoveNext
loop
RS1.Close
set RS1 = Nothing

arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No	&",,,,",",")
arrPSP_Count		= split(strPSP_Count		&",,,,",",")
arrPSP_ST			= split(strPSP_ST			&",,,,",",")
arrPSP_Desc			= split(strPSP_Desc			&",,,,",",")
arrPSP_Start		= split(strPSP_Start		&",,,,",",")
arrPSP_End			= split(strPSP_End			&",,,,",",")
arrPSP_Sub_YN		= split(strPSP_Sub_YN		&",,,,",",")
arrPSP_Sub_Start	= split(strPSP_Sub_Start	&",,,,",",")
arrPSP_Sub_End		= split(strPSP_Sub_End		&",,,,",",")

'�迭ȭ �����Ƿ� ���ڿ� ������ �ٸ� �뵵�� ����ϱ� ���� �ʱ�ȭ
strPSP_Sub_YN		= ""
strPSP_Sub_Start	= ""
strPSP_Sub_End		= ""

'���� PNO���� Merging �ع����� �����÷��װ� �������Ƿ�, ���� �����÷��� ������ ������ �����Ѵ�.
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
	if arrPSP_Sub_YN(CNT1) = "Y" and instr(strPSP_Sub_YN,arrBOM_Sub_BS_D_No(CNT1)) = 0 then
		strPSP_Sub_YN		= strPSP_Sub_YN		& arrBOM_Sub_BS_D_No(CNT1)	& ";"
		strPSP_Sub_Start	= strPSP_Sub_Start	& arrPSP_Sub_Start(CNT1)	& ";"
		strPSP_Sub_End		= strPSP_Sub_End	& arrPSP_Sub_End(CNT1)		& ";"
	end if
next

'���� ���� ���̾� ���� ��, Mergingó��
CNT3 = 0
for CNT1=1 to ubound(arrBOM_Sub_BS_D_No)-1
	'������ �𵨰� �����ϴٸ�
	if arrBOM_Sub_BS_D_No(CNT1) <> "" and arrBOM_Sub_BS_D_No(CNT1) = arrBOM_Sub_BS_D_No(CNT1-1) then
		'���� ���ڵ忡 ������ ���ϸ� �ȴ�.
		arrPSP_Count(CNT1-1)	= int(arrPSP_Count(CNT1-1)) + int(arrPSP_Count(CNT1))
		
		'���� �ִ� ĭ���� ��ĭ�� ��ܿ���
		for CNT2=CNT1 to ubound(arrBOM_Sub_BS_D_No)-2
			arrBOM_Sub_BS_D_No(CNT2)	= arrBOM_Sub_BS_D_No(CNT2+1)
			arrPSP_Count(CNT2)			= arrPSP_Count(CNT2+1)
			arrPSP_ST(CNT2)				= arrPSP_ST(CNT2+1)
			arrPSP_Desc(CNT2)			= arrPSP_Desc(CNT2+1)
			arrPSP_Start(CNT2)			= arrPSP_Start(CNT2+1)
			arrPSP_End(CNT2)			= arrPSP_End(CNT2+1)
		next
		arrBOM_Sub_BS_D_No(CNT2) = ""
		CNT1=CNT1-1 '���ڵ尡 �ϳ��� ����� ���̹Ƿ�, CNT1�� �ٽ� �����ϱ� ���� �����Ѵ�.
		CNT3=CNT3+1
	end if
next

'���ݱ����� ó�� ��Ȳ��, ��ȹ DB���� ����, Ư�������� ���ڵ带 �� �����ͼ�, �迭�� �����.
'���������� ��¥�� ��¡�� �Ͼ��, ���ǹ� �����Ƿ�, ���� ������ ��, ���߿� �ڹٽ�ũ��Ʈ�� ó�� �Ѵ�.
'�� ���� ��¡�ϸ鼭, �迭�� ũ�Ⱑ �پ� ��.

'������ ����, HTML�ϰ� ���İ��� �ϱ� ����, ���� �迭 ���� ����.
'OPT�迭�� �̰�
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)-1-CNT3	
	arrOptBOM_Sub_BS_D_No(CNT1)	= arrBOM_Sub_BS_D_No(CNT1)
	arrOptPSP_Count(CNT1)		= arrPSP_Count(CNT1)
	arrOptPSP_ST(CNT1)			= arrPSP_ST(CNT1)
	arrOptPSP_Desc(CNT1)		= arrPSP_Desc(CNT1)
	arrOptPSP_Start(CNT1)		= arrPSP_Start(CNT1)
	arrOptPSP_End(CNT1)			= arrPSP_End(CNT1)
next

'Start�� End �� ����� ���ؼ� �ٽ� ���Ҵ��� �ʿ���.
'�� �غ���.

'���� �޽������� �Է��Ѵ�.
'�޽� ���� �ð�(��)
arrRest(0,0) = 37200
arrRest(1,0) = 45000
arrRest(2,0) = 54600
arrRest(3,0) = 62400
'�޽� �ð�(��)
arrRest(0,1) = 600
arrRest(1,1) = 3000
arrRest(2,1) = 600
arrRest(3,1) = 1200


'���� �ð� �ʱ�ȭ 8�� 20���� �������� ���� 30000��°
nAccTime	= 30000

'�ð��� �ʴ����� ȯ�� �� ���� ���� �ð� �ʱ�ȭ
nRest = 0

for CNT1 = 0 to ubound(arrOptBOM_Sub_BS_D_No)
	if arrOptBOM_Sub_BS_D_No(CNT1) <> "" then
		
		'MC �ð� ��� ����
		if CNT1 = 0 then ' ù��° ���ڵ忡���� MC�� 0
			nMC_Time = 0	
		else
			nMC_Time = GetMCTime(oldBOM_Sub_BS_D_No, arrOptBOM_Sub_BS_D_No(CNT1)) * 60
		end if
		
		'���۽ð� ��, ����ð� ���, ����Ÿ�̸� ���� �� ����
		arrOptPSP_Start(CNT1)	= nAccTime	+ nMC_Time													'���� �ð��� ���� �ð� �ݿ� (MC�ð��� �ݿ�)
		arrOptPSP_End(CNT1)		= arrOptPSP_Start(CNT1) + (arrOptPSP_Count(CNT1) * arrOptPSP_ST(CNT1))	'����ð� ���
		nAccTime				= arrOptPSP_End(CNT1) + 1
		'���۽ð� ��, ����ð� ���, ����Ÿ�̸� ���� �� ��
		
		'��� �� ����ð��� ���� �ð��� ��ġ�� ��쿡 ���� ó�� ����
		for CNT2 = nRest to ubound(arrRest)	'�켱 ���� �ð� ��ŭ ����
			PS = int(arrOptPSP_Start(CNT1))						'��ȹ����
			PE = int(arrOptPSP_End(CNT1))						'��ȹ����
			RS = int(arrRest(CNT2,0))							'�޽Ľ���
			RE = int(arrRest(CNT2,0)) + int(arrRest(CNT2,1))	'�޽�����
			
			'��ȹ ���ᰡ �޽� ���۰� ���� �ɷ� �ִ� ��� or
			'��ȹ ������ �޽� ���۰� �� ���̿� �ɷ� �ִ� �ܿ� or
			'��ȹ ���۰� �� ���̿� �޽� �ð��� �ִ� ���
			if ((RS < PE and PE <= RE) or (RS <= PS and PS < RE) or (PS <= RS and RE <= PE)) then
				'���� �ð��� ��� �ΰ�, ���½ð� �ڷ� ���ڵ尡 �߰��ǹǷ� �ڷ� ��ĭ�� �̵�
				for CNT3 = ubound(arrOptBOM_Sub_BS_D_No)-2 to CNT1 step -1
					arrOptBOM_Sub_BS_D_No(CNT3+2)	= arrOptBOM_Sub_BS_D_No(CNT3+1)
					arrOptPSP_Count(CNT3+2)			= arrOptPSP_Count(CNT3+1)
					arrOptPSP_ST(CNT3+2)			= arrOptPSP_ST(CNT3+1)
					arrOptPSP_Start(CNT3+2)			= arrOptPSP_Start(CNT3+1)
					arrOptPSP_End(CNT3+2)			= arrOptPSP_End(CNT3+1)
					arrOptPSP_Desc(CNT3+2)			= arrOptPSP_Desc(CNT3+1)
				next				

				'����ROW�� ���� �� ROW �߰�.
				arrOptBOM_Sub_BS_D_No(CNT1+1)	= arrOptBOM_Sub_BS_D_No(CNT1)
				arrOptPSP_ST(CNT1+1)			= arrOptPSP_ST(CNT1)
				arrOptPSP_Desc(CNT1+1)			= arrOptPSP_Desc(CNT1)
				'������ ������ �Ҵ�
				
				if arrOptPSP_ST(CNT1) = 0 or isnull(arrOptPSP_ST(CNT1)) or arrOptPSP_ST(CNT1) = "" then
					arrOptPSP_Count(CNT1+1)	= arrOptPSP_Count(CNT1)
				else
					arrOptPSP_Count(CNT1+1)	= arrOptPSP_Count(CNT1) - Formatnumber((arrRest(CNT2,0) - int(arrOptPSP_Start(CNT1))) / arrOptPSP_ST(CNT1), 0)
					arrOptPSP_Count(CNT1)	= Formatnumber((arrRest(CNT2,0) - int(arrOptPSP_Start(CNT1))) / arrOptPSP_ST(CNT1), 0)
				end if
				
				'����ROW�� ���� �ð��� �޽� �������� ����
				arrOptPSP_End(CNT1)		= arrRest(CNT2,0)
				
				nRest	= nRest + 1 '�޽� ��ȣ ����

				'����ROW�� ���� �ð� ����
				nAccTime = arrRest(CNT2,0) + arrRest(CNT2,1) '�޽Ľð� ���� �ð�.
			end if
		next
		'��� �� ����ð��� ���� �ð��� ��ġ�� ��쿡 ���� ó�� ��
		
		arrOptPSP_Period(CNT1) = int(arrOptPSP_End(CNT1)-arrOptPSP_Start(CNT1))

		'���� �ð��� ���� �ð��� �ɸ���, ���� �ð� ���� �ð����� ����
		for CNT2 = 0 to ubound(arrRest)
			if nAccTime = arrRest(CNT2,0) then
				nAccTime = nAccTime + arrRest(CNT2,1)
				nRest	= nRest + 1 '�޽� ��ȣ ����
			end if
		next
	
		oldBOM_Sub_BS_D_No = arrOptBOM_Sub_BS_D_No(CNT1)	'MC�� �ݿ��� ���Ͽ�, ���� ��Ʈ�ѹ� ����.
	end if
next

call BOMSub_Guide()
%>

<html>
<head>
</head>
<body topmargin=0 leftmargin=0>

<script language="javascript">
function save_list_to_db(form)
{
	form.submit();
}

function insert_item(nCNT1)
{
	for(var i=frmPlan_State.BOM_Sub_BS_D_No.length-1; i > nCNT1; i--)
	{
		frmPlan_State.BOM_Sub_BS_D_No[i].value	= frmPlan_State.BOM_Sub_BS_D_No[i-1].value;
		frmPlan_State.PSP_Count[i].value		= frmPlan_State.PSP_Count[i-1].value;
		frmPlan_State.PSP_ST[i].value			= frmPlan_State.PSP_ST[i-1].value;
		frmPlan_State.PSP_Period[i].value		= frmPlan_State.PSP_Period[i-1].value;
		frmPlan_State.PSP_Start[i].value		= frmPlan_State.PSP_Start[i-1].value;
		frmPlan_State.PSP_End[i].value			= frmPlan_State.PSP_End[i-1].value;
		frmPlan_State.PSP_Desc[i].value			= frmPlan_State.PSP_Desc[i-1].value;
		frmPlan_State.PSP_Sub_YN[i].value		= frmPlan_State.PSP_Sub_YN[i-1].value;
		frmPlan_State.PSP_Sub_Start[i].value	= frmPlan_State.PSP_Sub_Start[i-1].value;
		frmPlan_State.PSP_Sub_End[i].value		= frmPlan_State.PSP_Sub_End[i-1].value;
	}
	frmPlan_State.BOM_Sub_BS_D_No[nCNT1].value	= "";
	frmPlan_State.PSP_Count[nCNT1].value		= "";
	frmPlan_State.PSP_ST[nCNT1].value			= "";
	frmPlan_State.PSP_Period[nCNT1].value		= "";
	frmPlan_State.PSP_Start[nCNT1].value		= "";
	frmPlan_State.PSP_End[nCNT1].value			= "";
	frmPlan_State.PSP_Desc[nCNT1].value			= "";
	frmPlan_State.PSP_Sub_YN[nCNT1].value		= "";
	frmPlan_State.PSP_Sub_Start[nCNT1].value	= "";
	frmPlan_State.PSP_Sub_End[nCNT1].value		= "";
}

function delete_item(nCNT1)
{
	for(var i=nCNT1; i < frmPlan_State.BOM_Sub_BS_D_No.length-1; i++)
	{
		i = parseInt(i);
		frmPlan_State.BOM_Sub_BS_D_No[i].value	= frmPlan_State.BOM_Sub_BS_D_No[i+1].value;
		frmPlan_State.PSP_Count[i].value		= frmPlan_State.PSP_Count[i+1].value;
		frmPlan_State.PSP_ST[i].value			= frmPlan_State.PSP_ST[i+1].value;
		frmPlan_State.PSP_Period[i].value		= frmPlan_State.PSP_Period[i+1].value;
		frmPlan_State.PSP_Start[i].value		= frmPlan_State.PSP_Start[i+1].value;
		frmPlan_State.PSP_End[i].value			= frmPlan_State.PSP_End[i+1].value;
		frmPlan_State.PSP_Desc[i].value			= frmPlan_State.PSP_Desc[i+1].value;
		frmPlan_State.PSP_Sub_YN[i].value		= frmPlan_State.PSP_Sub_YN[i+1].value;
		frmPlan_State.PSP_Sub_Start[i].value	= frmPlan_State.PSP_Sub_Start[i+1].value;
		frmPlan_State.PSP_Sub_End[i].value		= frmPlan_State.PSP_Sub_End[i+1].value;
	}
	frmPlan_State.BOM_Sub_BS_D_No[i].value	= "";
	frmPlan_State.PSP_Count[i].value		= "";
	frmPlan_State.PSP_ST[i].value			= "";
	frmPlan_State.PSP_Period[i].value		= "";
	frmPlan_State.PSP_Start[i].value		= "";
	frmPlan_State.PSP_End[i].value			= "";
	frmPlan_State.PSP_Desc[i].value			= "";
	frmPlan_State.PSP_Sub_YN[i].value		= "";
	frmPlan_State.PSP_Sub_Start[i].value	= "";
	frmPlan_State.PSP_Sub_End[i].value		= "";
}
</script>

<table width=420px cellpadding=0 cellspacing=1>
<form name="frmPlan_State_list" action="Process_State_plan_action.asp" method="post">
<input type="hidden" name="s_Work_Date"			value="<%=request("s_Work_Date")%>">
<input type="hidden" name="s_Line"				value="<%=request("s_Line")%>">
<tr height=400px>
	<td id="idPlan_State">
		<table width=100% border=1 height=400px>
		<tr>
			<td width=40px>�۾�<br>��</td>
			<td width=100px><textarea name="lstBOM_Sub_BS_D_No" cols=15 style="height:100%;"></textarea></td>
			<td width=40px>��ǥ<br>����</td>
			<td width=100px><textarea name="lstPSP_Count" cols=15 style="height:100%;"></textarea></td>
			<td width=40px>T.T(s)</td>
			<td width=100px><textarea name="lstPSP_ST" cols=15 style="height:100%;"></textarea></td>
		</tr>
	</td>
</tr>
<tr height=22px>
	<td colspan=6><input type="button" value="������ ����" onclick="javascript:save_list_to_db(this.form)"></td>
</tr>
</form>
</table>

<br><br><br><br>
<table width=675px cellpadding=0 cellspacing=1>
<form name="frmPlan_State" action="Process_State_plan_action.asp" method="post">
<input type="hidden" name="s_Work_Date"			value="<%=request("s_Work_Date")%>">
<input type="hidden" name="s_Line"				value="<%=request("s_Line")%>">
<tr>
	<td><input type="button" value="������ ����" onclick="javascript:save_list_to_db(this.form)"></td>
</tr>
<tr>
	<td id="idPlan_State">
		<table width=100% border=1>
		<tr>
			<td width=90px>�۾���</td>
			<td width=40px>��ǥ<br>����</td>
			<td width=40px>T.T(s)</td>
			<td width=55px>����<br>�ð�(m)</td>
			<td width=55px>����</td>
			<td width=55px>����</td>
			<td width=130px>���</td>
			<td width=120px>�����۾�</td>
			<td width=90px>�۾�</td>
		</tr>
<%
'�� 60���� ����
for CNT1 = 0 to ubound(arrOptBOM_Sub_BS_D_No)
	'���� �ʱ�ȭ	
	PSP_Period		= ""
	PSP_Start	= ""
	strPSP_End		= ""
	
	'DB���� ������ ������ ������ ����
	
	'�� ǥ�ø� �ð�:�� ǥ�÷� ����
	PSP_Period		= ""
	strPSP_Start	= ""
	strPSP_End		= ""
		
	if arrOptBOM_Sub_BS_D_No(CNT1) <> ""  then	'DBũ�⸸ŭ 
		PSP_Period	= formatnumber(arrOptPSP_Period(CNT1) / 60,0)
		PSP_Start	= int(arrOptPSP_Start(CNT1) / 60)
		PSP_End		= int(arrOptPSP_End(CNT1) / 60)
		
		if INT(PSP_Start / 60) < 10 then
			strPSP_Start = strPSP_Start & "0"
		end if
		strPSP_Start = strPSP_Start & CSTR(INT(PSP_Start / 60))
		if INT(PSP_Start mod 60) < 10 then
			strPSP_Start = strPSP_Start & "0"
		end if
		strPSP_Start = strPSP_Start & CSTR(PSP_Start mod 60)
		
		if INT(PSP_End / 60) < 10 then
			strPSP_End = strPSP_End & "0"
		end if
		strPSP_End = strPSP_End & CSTR(INT(PSP_End / 60))
		if INT(PSP_End mod 60) < 10 then
			strPSP_End = strPSP_End & "0"
		end if
		strPSP_End = strPSP_End & CSTR(PSP_End mod 60)
	end if
%>
		<tr style="height:10px">
			<td><input type='text' name='BOM_Sub_BS_D_No'	style='width:98%'	readonly	value="<%=arrOptBOM_Sub_BS_D_No(CNT1)%>" onclick='javascript:show_BOMSub_Guide(this);'></td>
			<td><input type='text' name='PSP_Count'			style='width:98%'				value="<%=arrOptPSP_Count(CNT1)%>"></td>
			<td><input type='text' name='PSP_ST'			style='width:98%'				value="<%=arrOptPSP_ST(CNT1)%>"></td>
			<td><input type='text' name='PSP_Period'		style='width:98%'	readonly	value="<%=PSP_Period%>"></td>
			<td><input type='text' name='PSP_Start'			style='width:98%'	readonly	value="<%=strPSP_Start%>"></td>
			<td><input type='text' name='PSP_End'			style='width:98%'	readonly	value="<%=strPSP_End%>"></td>
			<td><input type='text' name='PSP_Desc'			style='width:98%'				value="<%=arrOptPSP_Desc(CNT1)%>"></td>
			<td>
				<input type='checkbox'	name='PSP_Sub_YN' 		value="<%=CNT1%>">
				<input type='text'		name='PSP_Sub_Start'	value=""	maxlength=4 size=4>
				<input type='text'		name='PSP_Sub_End'		value=""	maxlength=4 size=4>
			</td>
			<td>
				<input type='button' value='����' onclick="javascript:insert_item('<%=CNT1%>')">
				<input type='button' value='����' onDblclick="javascript:delete_item('<%=CNT1%>')">
			</td>
		</tr>
<%
next
%>
		</table>
	</td>
</tr>
<tr>
	<td><input type="button" value="������ ����" onclick="javascript:save_list_to_db(this.form)"></td>
</tr>
</form>
</table>
</body>
</html>

<script language="javascript">
	//�����÷��� ���� �迭ȭ
	var arrPSP_Sub_YN		= "<%=strPSP_Sub_YN%>".split(";");
	var arrPSP_Sub_Start	= "<%=strPSP_Sub_Start%>".split(";");
	var arrPSP_Sub_End		= "<%=strPSP_Sub_End%>".split(";");
	
	for (var i=0; i < arrPSP_Sub_YN.length-1; i++)
	{
		for (var j=0; j < frmPlan_State.BOM_Sub_BS_D_No.length; j++)
		{
			//���̺�� �����÷��� ���� ���Ͽ� �Ҵ� �Ѵ�.
			if (arrPSP_Sub_YN[i] != "" && frmPlan_State.BOM_Sub_BS_D_No[j].value == arrPSP_Sub_YN[i])
			{
				frmPlan_State.PSP_Sub_YN[j].checked		= true;
				frmPlan_State.PSP_Sub_Start[j].value	= arrPSP_Sub_Start[i];
				frmPlan_State.PSP_Sub_End[j].value		= arrPSP_Sub_End[i];
			}
		}
	}
	
	//�ð� ����� �ȵ� ���¶��, �ٽ� �����Ѵ�.
	var ReSave_Require_YN = "<%=request("s_ReSave_Require_YN")%>";
	if (ReSave_Require_YN == "Y")
		save_list_to_db(frmPlan_State);
</script>

<%
function GetMCTime(strBS_D_No1, strBS_D_No2)
	if left(strBS_D_No1,5) = "6871A" then
		if left(strBS_D_No1,10) <> left(strBS_D_No2,10) then
			GetMCTime = 3
		end if
	else
		if left(strBS_D_No1,9) <> left(strBS_D_No2,9) then
			GetMCTime = 3
		end if
	end if
end function
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->