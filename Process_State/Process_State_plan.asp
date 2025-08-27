<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->


<%
'DB용 변수 선언
dim RS1
dim SQL

'반복문에 사용하기 위한 변수 선언
dim CNT1
dim CNT2
dim CNT3


dim strBOM_Sub_BS_D_No	'파트넘버
dim strPSP_Count		'목표수량
dim strPSP_ST			'표준시간
dim strPSP_Desc			'비고
dim strPSP_Start		'시작시각
dim strPSP_End			'종료시각

dim arrBOM_Sub_BS_D_No	
dim arrPSP_Count		
dim arrPSP_ST			
dim arrPSP_Desc			
dim arrPSP_Start		
dim arrPSP_End			


dim strPSP_Sub_YN		'서브PCB 플래그
dim strPSP_Sub_Start	'시작시각
dim strPSP_Sub_End		'종료시각

dim arrPSP_Sub_YN		
dim arrPSP_Sub_Start	
dim arrPSP_Sub_End		

'최적화 된 배열
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

'모델체인지 시간
dim nMC_Time

'소요시간
dim PSP_Period

'유효 행 길이
dim nDBRowLength

'누적된 시간표시
dim nAccTime

'누적된 쉬는 시간 표시
dim nRest

'쉬는 시간 데이타 할당
dim arrRest(3,1)

dim PS	'계획 시작
dim PE	'계획 종료
dim RS	'휴식 시작
dim RE	'휴식 종료

'계획을 DB에서 조회, 라인과 날짜로 조회
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

'배열화 했으므로 문자열 변수는 다른 용도로 사용하기 위해 초기화
strPSP_Sub_YN		= ""
strPSP_Sub_Start	= ""
strPSP_Sub_End		= ""

'같은 PNO끼리 Merging 해버리면 서브플래그가 지워지므로, 지금 서브플래그 정보를 변수에 저장한다.
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
	if arrPSP_Sub_YN(CNT1) = "Y" and instr(strPSP_Sub_YN,arrBOM_Sub_BS_D_No(CNT1)) = 0 then
		strPSP_Sub_YN		= strPSP_Sub_YN		& arrBOM_Sub_BS_D_No(CNT1)	& ";"
		strPSP_Sub_Start	= strPSP_Sub_Start	& arrPSP_Sub_Start(CNT1)	& ";"
		strPSP_Sub_End		= strPSP_Sub_End	& arrPSP_Sub_End(CNT1)		& ";"
	end if
next

'동일 모델이 연이어 있을 때, Merging처리
CNT3 = 0
for CNT1=1 to ubound(arrBOM_Sub_BS_D_No)-1
	'직전의 모델과 동일하다면
	if arrBOM_Sub_BS_D_No(CNT1) <> "" and arrBOM_Sub_BS_D_No(CNT1) = arrBOM_Sub_BS_D_No(CNT1-1) then
		'직전 레코드에 수량만 합하면 된다.
		arrPSP_Count(CNT1-1)	= int(arrPSP_Count(CNT1-1)) + int(arrPSP_Count(CNT1))
		
		'지금 있는 칸으로 한칸씩 당겨오기
		for CNT2=CNT1 to ubound(arrBOM_Sub_BS_D_No)-2
			arrBOM_Sub_BS_D_No(CNT2)	= arrBOM_Sub_BS_D_No(CNT2+1)
			arrPSP_Count(CNT2)			= arrPSP_Count(CNT2+1)
			arrPSP_ST(CNT2)				= arrPSP_ST(CNT2+1)
			arrPSP_Desc(CNT2)			= arrPSP_Desc(CNT2+1)
			arrPSP_Start(CNT2)			= arrPSP_Start(CNT2+1)
			arrPSP_End(CNT2)			= arrPSP_End(CNT2+1)
		next
		arrBOM_Sub_BS_D_No(CNT2) = ""
		CNT1=CNT1-1 '레코드가 하나씩 당겨진 셈이므로, CNT1을 다시 수행하기 위해 차감한다.
		CNT3=CNT3+1
	end if
next

'지금까지의 처리 상황은, 계획 DB에서 금일, 특정라인의 레코드를 쭉 가져와서, 배열에 담았음.
'서브정보는 어짜피 머징이 일어나면, 무의미 해지므로, 따로 정리해 둠, 나중에 자바스크립트로 처리 한다.
'그 다음 머징하면서, 배열의 크기가 줄어 듬.

'앞으로 할일, HTML하고 스파게팅 하기 전에, 먼저 배열 정리 하자.
'OPT배열에 이관
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)-1-CNT3	
	arrOptBOM_Sub_BS_D_No(CNT1)	= arrBOM_Sub_BS_D_No(CNT1)
	arrOptPSP_Count(CNT1)		= arrPSP_Count(CNT1)
	arrOptPSP_ST(CNT1)			= arrPSP_ST(CNT1)
	arrOptPSP_Desc(CNT1)		= arrPSP_Desc(CNT1)
	arrOptPSP_Start(CNT1)		= arrPSP_Start(CNT1)
	arrOptPSP_End(CNT1)			= arrPSP_End(CNT1)
next

'Start와 End 는 계산을 통해서 다시 재할당이 필요함.
'자 해보자.

'먼저 휴식정보를 입력한다.
'휴식 시작 시간(초)
arrRest(0,0) = 37200
arrRest(1,0) = 45000
arrRest(2,0) = 54600
arrRest(3,0) = 62400
'휴식 시간(초)
arrRest(0,1) = 600
arrRest(1,1) = 3000
arrRest(2,1) = 600
arrRest(3,1) = 1200


'누적 시간 초기화 8시 20분이 자정으로 부터 30000초째
nAccTime	= 30000

'시간을 초단위로 환산 및 누적 쉬는 시간 초기화
nRest = 0

for CNT1 = 0 to ubound(arrOptBOM_Sub_BS_D_No)
	if arrOptBOM_Sub_BS_D_No(CNT1) <> "" then
		
		'MC 시간 계산 시작
		if CNT1 = 0 then ' 첫번째 레코드에서는 MC는 0
			nMC_Time = 0	
		else
			nMC_Time = GetMCTime(oldBOM_Sub_BS_D_No, arrOptBOM_Sub_BS_D_No(CNT1)) * 60
		end if
		
		'시작시간 및, 종료시간 계산, 누적타이머 증가 등 시작
		arrOptPSP_Start(CNT1)	= nAccTime	+ nMC_Time													'시작 시간에 누적 시간 반영 (MC시간도 반영)
		arrOptPSP_End(CNT1)		= arrOptPSP_Start(CNT1) + (arrOptPSP_Count(CNT1) * arrOptPSP_ST(CNT1))	'종료시간 계산
		nAccTime				= arrOptPSP_End(CNT1) + 1
		'시작시간 및, 종료시간 계산, 누적타이머 증가 등 끝
		
		'계산 된 종료시간이 쉬는 시간에 겹치는 경우에 대한 처리 시작
		for CNT2 = nRest to ubound(arrRest)	'우선 쉬는 시간 만큼 루프
			PS = int(arrOptPSP_Start(CNT1))						'계획시작
			PE = int(arrOptPSP_End(CNT1))						'계획종료
			RS = int(arrRest(CNT2,0))							'휴식시작
			RE = int(arrRest(CNT2,0)) + int(arrRest(CNT2,1))	'휴식종료
			
			'계획 종료가 휴식 시작과 끝에 걸려 있는 경우 or
			'계획 시작이 휴식 시작과 끝 사이에 걸려 있는 겨우 or
			'계획 시작과 끝 사이에 휴식 시간이 있는 경우
			if ((RS < PE and PE <= RE) or (RS <= PS and PS < RE) or (PS <= RS and RE <= PE)) then
				'쉬는 시간을 가운데 두고, 쉬는시간 뒤로 레코드가 추가되므로 뒤로 한칸씩 이동
				for CNT3 = ubound(arrOptBOM_Sub_BS_D_No)-2 to CNT1 step -1
					arrOptBOM_Sub_BS_D_No(CNT3+2)	= arrOptBOM_Sub_BS_D_No(CNT3+1)
					arrOptPSP_Count(CNT3+2)			= arrOptPSP_Count(CNT3+1)
					arrOptPSP_ST(CNT3+2)			= arrOptPSP_ST(CNT3+1)
					arrOptPSP_Start(CNT3+2)			= arrOptPSP_Start(CNT3+1)
					arrOptPSP_End(CNT3+2)			= arrOptPSP_End(CNT3+1)
					arrOptPSP_Desc(CNT3+2)			= arrOptPSP_Desc(CNT3+1)
				next				

				'다음ROW에 동일 모델 ROW 추가.
				arrOptBOM_Sub_BS_D_No(CNT1+1)	= arrOptBOM_Sub_BS_D_No(CNT1)
				arrOptPSP_ST(CNT1+1)			= arrOptPSP_ST(CNT1)
				arrOptPSP_Desc(CNT1+1)			= arrOptPSP_Desc(CNT1)
				'수량을 나누어 할당
				
				if arrOptPSP_ST(CNT1) = 0 or isnull(arrOptPSP_ST(CNT1)) or arrOptPSP_ST(CNT1) = "" then
					arrOptPSP_Count(CNT1+1)	= arrOptPSP_Count(CNT1)
				else
					arrOptPSP_Count(CNT1+1)	= arrOptPSP_Count(CNT1) - Formatnumber((arrRest(CNT2,0) - int(arrOptPSP_Start(CNT1))) / arrOptPSP_ST(CNT1), 0)
					arrOptPSP_Count(CNT1)	= Formatnumber((arrRest(CNT2,0) - int(arrOptPSP_Start(CNT1))) / arrOptPSP_ST(CNT1), 0)
				end if
				
				'현재ROW의 종료 시각을 휴식 시작으로 조정
				arrOptPSP_End(CNT1)		= arrRest(CNT2,0)
				
				nRest	= nRest + 1 '휴식 번호 증가

				'다음ROW의 시작 시각 조정
				nAccTime = arrRest(CNT2,0) + arrRest(CNT2,1) '휴식시간 종료 시각.
			end if
		next
		'계산 된 종료시간이 쉬는 시간에 겹치는 경우에 대한 처리 끝
		
		arrOptPSP_Period(CNT1) = int(arrOptPSP_End(CNT1)-arrOptPSP_Start(CNT1))

		'누적 시간이 쉬는 시간에 걸리면, 쉬는 시간 종료 시간으로 변경
		for CNT2 = 0 to ubound(arrRest)
			if nAccTime = arrRest(CNT2,0) then
				nAccTime = nAccTime + arrRest(CNT2,1)
				nRest	= nRest + 1 '휴식 번호 증가
			end if
		next
	
		oldBOM_Sub_BS_D_No = arrOptBOM_Sub_BS_D_No(CNT1)	'MC를 반영을 위하여, 직전 파트넘버 저장.
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
			<td width=40px>작업<br>모델</td>
			<td width=100px><textarea name="lstBOM_Sub_BS_D_No" cols=15 style="height:100%;"></textarea></td>
			<td width=40px>목표<br>수량</td>
			<td width=100px><textarea name="lstPSP_Count" cols=15 style="height:100%;"></textarea></td>
			<td width=40px>T.T(s)</td>
			<td width=100px><textarea name="lstPSP_ST" cols=15 style="height:100%;"></textarea></td>
		</tr>
	</td>
</tr>
<tr height=22px>
	<td colspan=6><input type="button" value="현재목록 저장" onclick="javascript:save_list_to_db(this.form)"></td>
</tr>
</form>
</table>

<br><br><br><br>
<table width=675px cellpadding=0 cellspacing=1>
<form name="frmPlan_State" action="Process_State_plan_action.asp" method="post">
<input type="hidden" name="s_Work_Date"			value="<%=request("s_Work_Date")%>">
<input type="hidden" name="s_Line"				value="<%=request("s_Line")%>">
<tr>
	<td><input type="button" value="현재목록 저장" onclick="javascript:save_list_to_db(this.form)"></td>
</tr>
<tr>
	<td id="idPlan_State">
		<table width=100% border=1>
		<tr>
			<td width=90px>작업모델</td>
			<td width=40px>목표<br>수량</td>
			<td width=40px>T.T(s)</td>
			<td width=55px>예상<br>시간(m)</td>
			<td width=55px>시작</td>
			<td width=55px>종료</td>
			<td width=130px>비고</td>
			<td width=120px>서브작업</td>
			<td width=90px>작업</td>
		</tr>
<%
'총 60행을 만듬
for CNT1 = 0 to ubound(arrOptBOM_Sub_BS_D_No)
	'변수 초기화	
	PSP_Period		= ""
	PSP_Start	= ""
	strPSP_End		= ""
	
	'DB에서 가져온 값들은 변수에 넣음
	
	'분 표시를 시간:분 표시로 변경
	PSP_Period		= ""
	strPSP_Start	= ""
	strPSP_End		= ""
		
	if arrOptBOM_Sub_BS_D_No(CNT1) <> ""  then	'DB크기만큼 
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
				<input type='button' value='삽입' onclick="javascript:insert_item('<%=CNT1%>')">
				<input type='button' value='삭제' onDblclick="javascript:delete_item('<%=CNT1%>')">
			</td>
		</tr>
<%
next
%>
		</table>
	</td>
</tr>
<tr>
	<td><input type="button" value="현재목록 저장" onclick="javascript:save_list_to_db(this.form)"></td>
</tr>
</form>
</table>
</body>
</html>

<script language="javascript">
	//서브플래그 정보 배열화
	var arrPSP_Sub_YN		= "<%=strPSP_Sub_YN%>".split(";");
	var arrPSP_Sub_Start	= "<%=strPSP_Sub_Start%>".split(";");
	var arrPSP_Sub_End		= "<%=strPSP_Sub_End%>".split(";");
	
	for (var i=0; i < arrPSP_Sub_YN.length-1; i++)
	{
		for (var j=0; j < frmPlan_State.BOM_Sub_BS_D_No.length; j++)
		{
			//테이블과 서브플래그 값을 비교하여 할당 한다.
			if (arrPSP_Sub_YN[i] != "" && frmPlan_State.BOM_Sub_BS_D_No[j].value == arrPSP_Sub_YN[i])
			{
				frmPlan_State.PSP_Sub_YN[j].checked		= true;
				frmPlan_State.PSP_Sub_Start[j].value	= arrPSP_Sub_Start[i];
				frmPlan_State.PSP_Sub_End[j].value		= arrPSP_Sub_End[i];
			}
		}
	}
	
	//시간 계산이 안된 상태라면, 다시 저장한다.
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