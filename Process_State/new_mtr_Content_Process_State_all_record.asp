<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
response.write now()
'변수 선언
dim RS1
dim RS2
dim SQL

'목표 수량을 계산하기 위한
dim sumPRD_Code '실적 수량 합
dim BS_ST		'ST 저장용

'현재 시간을 계산하기 위한 변수, 라인별 선언
dim calcNow
dim calcNow1
dim calcNow2
dim calcNow3
dim calcNow4
dim calcNow5
dim calcNow6
dim calcNow7

dim cntPRD_Code

'투입 집계용 변수
dim isum1
dim isum2
dim isum3
dim isum4
dim isum5
dim isum6
dim isum7

'실적 집계용 변수
dim sum1
dim sum2
dim sum3
dim sum4
dim sum5
dim sum6
dim sum7

'계획 집계용 변수
dim psum1
dim psum2
dim psum3
dim psum4
dim psum5
dim psum6
dim psum7

dim tsum1
dim tsum2
dim tsum3
dim tsum4
dim tsum5
dim tsum6
dim tsum7

'달성률 계산용 변수
dim rate1
dim rate2
dim rate3
dim rate4
dim rate5
dim rate6
dim rate7
dim rateSum

'배경색 저장용 변수
dim strBgClr1
dim strBgClr2
dim strBgClr3
dim strBgClr4
dim strBgClr5
dim strBgClr6
dim strBgClr7
dim strBgClrSum

dim strTRBgClr1
dim strTRBgClr2
dim strTRBgClr3
dim strTRBgClr4
dim strTRBgClr5
dim strTRBgClr6
dim strTRBgClr7
dim strTRBgClrSum

dim strLineState1
dim strLineState2
dim strLineState3
dim strLineState4
dim strLineState5
dim strLineState6
dim strLineState7

'날짜 변수
dim s_Work_Date
dim s_Process

'SQL = "insert into tbTest_setinterval (ts_Work,ts_Desc,ts_Now,ts_Diff) values ('ProcessState','ALL',getdate(),0)"
'sys_DBCon.execute(SQL)

'날짜는 무조건 현재 날짜로
's_Work_Date = request("s_Work_Date")
s_Work_Date = date()
s_Process = request("s_Process")

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

'실적 초기화
sum1 = 0
sum2 = 0
sum3 = 0
sum4 = 0
sum5 = 0
sum6 = 0
sum7 = 0

'투입때문에 추가
isum1 = 0
isum2 = 0
isum3 = 0
isum4 = 0
isum5 = 0
isum6 = 0
isum7 = 0

'라인별로, 포장 완료 수량을 집계
SQL = "select PRD_Line, cntPRD_Code = count(PRD_Code) from tbPWS_Raw_Data "
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and "&vbcrlf
SQL = SQL & "	PRD_BOX_Date <> '' and PRD_BOX_Date is not null and PRD_Dummy_YN is null "&vbcrlf
SQL = SQL & "group by PRD_Line"&vbcrlf
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	if s_Process = "PCBA" then 
		if ucase(RS1("PRD_Line")) = "PCBA1" then
			sum1 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA2" then
			sum2 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA3" then
			sum3 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA4" then
			sum4 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA5" then
			sum5 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA6" then
			sum6 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA7" then
			sum7 = RS1("cntPRD_Code")
		end if					
	elseif s_Process = "CBOX" then 
		if ucase(RS1("PRD_Line")) = "CBOX1" then
			sum1 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX2" then
			sum2 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX3" then
			sum3 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX4" then
			sum4 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX5" then
			sum5 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX6" then
			sum6 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX7" then
			sum7 = RS1("cntPRD_Code")
		end if	
	end if
	RS1.MoveNext
loop
RS1.Close

'투입때문에 추가
SQL = "select PRD_Line, cntPRD_Code = count(PRD_Code) from tbPWS_Raw_Data "
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and PRD_Dummy_YN is null "&vbcrlf
'SQL = SQL & "	PRD_BOX_Date = '"&s_Work_Date&"' "&vbcrlf
SQL = SQL & "group by PRD_Line"&vbcrlf
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	if s_Process = "PCBA" then 
		if ucase(RS1("PRD_Line")) = "PCBA1" then
			isum1 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA2" then
			isum2 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA3" then
			isum3 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA4" then
			isum4 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA5" then
			isum5 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA6" then
			isum6 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "PCBA7" then
			isum7 = RS1("cntPRD_Code")
		end if					
	elseif s_Process = "CBOX" then 
		if ucase(RS1("PRD_Line")) = "CBOX1" then
			isum1 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX2" then
			isum2 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX3" then
			isum3 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX4" then
			isum4 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX5" then
			isum5 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX6" then
			isum6 = RS1("cntPRD_Code")
		elseif ucase(RS1("PRD_Line")) = "CBOX7" then
			isum7 = RS1("cntPRD_Code")
		end if	
	end if
	RS1.MoveNext
loop
RS1.Close

'서브 계획이 있다면, 실적으로 잡아서 추가한다.
SQL = "select PSP_Line, sumPSP_Count = sum(PSP_Count) from tbProcess_State_Plan "
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PSP_Sub_YN = 'Y' and len(PSP_Sub_Start) = 4 and len(PSP_Sub_End) = 4 and "&vbcrlf
SQL = SQL & "	PSP_Work_Date = '"&s_Work_Date&"' "
SQL = SQL & "group by PSP_Line"&vbcrlf
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	if s_Process = "PCBA" then 
		if ucase(RS1("PSP_Line")) = "PCBA1" then
			sum1 = sum1 + RS1("sumPSP_Count")
			'투입때문에 추가
			isum1 = isum1 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "PCBA2" then
			sum2 = sum2 + RS1("sumPSP_Count")
			isum2 = isum2 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "PCBA3" then
			sum3 = sum3 + RS1("sumPSP_Count")
			isum3 = isum3 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "PCBA4" then
			sum4 = sum4 + RS1("sumPSP_Count")
			isum4 = isum4 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "PCBA5" then
			sum5 = sum5 + RS1("sumPSP_Count")
			isum5 = isum5 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "PCBA6" then
			sum6 = sum6 + RS1("sumPSP_Count")
			isum6 = isum6 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "PCBA7" then
			sum7 = sum7 + RS1("sumPSP_Count")
			isum7 = isum7 + RS1("sumPSP_Count")
		end if				
	elseif s_Process="CBOX" then
		if ucase(RS1("PSP_Line")) = "CBOX1" then
			sum1 = sum1 + RS1("sumPSP_Count")
			isum1 = isum1 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "CBOX2" then
			sum2 = sum2 + RS1("sumPSP_Count")
			isum2 = isum2 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "CBOX3" then
			sum3 = sum3 + RS1("sumPSP_Count")
			isum3 = isum3 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "CBOX4" then
			sum4 = sum4 + RS1("sumPSP_Count")
			isum4 = isum4 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "CBOX5" then
			sum5 = sum5 + RS1("sumPSP_Count")
			isum5 = isum5 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "CBOX6" then
			sum6 = sum6 + RS1("sumPSP_Count")
			isum6 = isum6 + RS1("sumPSP_Count")
		elseif ucase(RS1("PSP_Line")) = "CBOX7" then
			sum7 = sum7 + RS1("sumPSP_Count")
			isum7 = isum7 + RS1("sumPSP_Count")
		end if		
	end if	
	RS1.MoveNext
loop
RS1.Close

'목표 수량을 분석하기 위해서
'지금 시각을 기준으로 자정으로 부터 몇 분이나 흘렀는지 확인
calcNow = left(FormatDateTime(now(),4),2)*60 + right(FormatDateTime(now(),4),2)

'쉬는 시간 중이라면, 쉬는 시간 시작 상태로 고정
if calcNow > 620 and calcNow <= 630 then
	calcNow = 620
end if
if calcNow > 750 and calcNow <= 790 then
	calcNow = 750
end if
if calcNow > 910 and calcNow <= 920 then
	calcNow = 910
end if

if calcNow > 1040 and calcNow <= 1060 then
	calcNow = 1040
end if

'쉬는 시간을 거친 수 만큼 쉬는 시간 차감
if calcNow > 630 then '10시 30분
	calcNow = calcNow - 10
end if
if calcNow > 790 then '13시 10분
	calcNow = calcNow - 40
end if
if calcNow > 920 then '15시 20분
	calcNow = calcNow - 10
end if
if calcNow > 1060 then '17시 40분
	calcNow = calcNow - 20
end if

calcNow = calcNow * 60

'각 라인별로 사용하기 위해서 변수별 분배 및 초로 환산
calcNow1 = calcNow
calcNow2 = calcNow
calcNow3 = calcNow
calcNow4 = calcNow
calcNow5 = calcNow
calcNow6 = calcNow
calcNow7 = calcNow

'각 라인별 최초 생산 시간으로 차감
SQL = "select PRD_Line, minPRD_Input_Time = min(PRD_Input_Time) from tbPWS_Raw_Data "
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and "&vbcrlf
SQL = SQL & "	PRD_Input_Date is not null "&vbcrlf
SQL = SQL & "group by PRD_Line"&vbcrlf

RS1.Open SQL,sys_DBCon
do until RS1.Eof
	if s_Process = "PCBA" then
		if ucase(RS1("PRD_Line")) = "PCBA1" then
			calcNow1 = calcNow1 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "PCBA2" then
			calcNow2 = calcNow2 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "PCBA3" then
			calcNow3 = calcNow3 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "PCBA4" then
			calcNow4 = calcNow4 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "PCBA5" then
			calcNow5 = calcNow5 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "PCBA6" then
			calcNow6 = calcNow6 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "PCBA7" then
			calcNow7 = calcNow7 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		end if
	elseif s_Process = "CBOX" then
		if ucase(RS1("PRD_Line")) = "CBOX1" then
			calcNow1 = calcNow1 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "CBOX2" then
			calcNow2 = calcNow2 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "CBOX3" then
			calcNow3 = calcNow3 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "CBOX4" then
			calcNow4 = calcNow4 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "CBOX5" then
			calcNow5 = calcNow5 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "CBOX6" then
			calcNow6 = calcNow6 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		elseif ucase(RS1("PRD_Line")) = "CBOX7" then
			calcNow7 = calcNow7 - (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		end if
	end if					
	RS1.MoveNext
loop
RS1.Close

'각 라인별 최초 생산 시간으로 차감이 안되었다면...
if calcNow = calcNow1 then
	calcNow1 = 0
end if
if calcNow = calcNow2 then
	calcNow2 = 0
end if
if calcNow = calcNow3 then
	calcNow3 = 0
end if
if calcNow = calcNow4 then
	calcNow4 = 0
end if
if calcNow = calcNow5 then
	calcNow5 = 0
end if
if calcNow = calcNow6 then
	calcNow6 = 0
end if
if calcNow = calcNow7 then
	calcNow7 = 0
end if

SQL = "select * from tbLine_State where left(LS_Line,4) = '"&s_Process&"' Order by LS_Line asc"
RS1.Open SQL,sys_DBCon
strLineState1 = RS1("LS_State")
if strLineState1 = "정상" then
	strBgClr1 = "black"
elseif strLineState1 = "품질" then
	strBgClr1 = "red"
elseif strLineState1 = "품절" then
	strBgClr1 = "green"
elseif strLineState1 = "설비" then
	strBgClr1 = "blue"
end if
RS1.MoveNext
strLineState2 = RS1("LS_State")
if strLineState2 = "정상" then
	strBgClr2 = "black"
elseif strLineState2 = "품질" then
	strBgClr2 = "red"
elseif strLineState2 = "품절" then
	strBgClr2 = "green"
elseif strLineState2 = "설비" then
	strBgClr2 = "blue"
end if
RS1.MoveNext
strLineState3 = RS1("LS_State")
if strLineState3 = "정상" then
	strBgClr3 = "black"
elseif strLineState3 = "품질" then
	strBgClr3 = "red"
elseif strLineState3 = "품절" then
	strBgClr3 = "green"
elseif strLineState3 = "설비" then
	strBgClr3 = "blue"
end if
RS1.MoveNext
strLineState4 = RS1("LS_State")
if strLineState4 = "정상" then
	strBgClr4 = "black"
elseif strLineState4 = "품질" then
	strBgClr4 = "red"
elseif strLineState4 = "품절" then
	strBgClr4 = "green"
elseif strLineState4 = "설비" then
	strBgClr4 = "blue"
end if
RS1.MoveNext
strLineState5 = RS1("LS_State")
if strLineState5 = "정상" then
	strBgClr5 = "black"
elseif strLineState5 = "품질" then
	strBgClr5 = "red"
elseif strLineState5 = "품절" then
	strBgClr5 = "green"
elseif strLineState5 = "설비" then
	strBgClr5 = "blue"
end if
RS1.MoveNext
strLineState6 = RS1("LS_State")
if strLineState6 = "정상" then
	strBgClr6 = "black"
elseif strLineState6 = "품질" then
	strBgClr6 = "red"
elseif strLineState6 = "품절" then
	strBgClr6 = "green"
elseif strLineState6 = "설비" then
	strBgClr6 = "blue"
end if
RS1.MoveNext
strLineState7 = RS1("LS_State")
if strLineState7 = "정상" then
	strBgClr7 = "black"
elseif strLineState7 = "품질" then
	strBgClr7 = "red"
elseif strLineState7 = "품절" then
	strBgClr7 = "green"
elseif strLineState7 = "설비" then
	strBgClr7 = "blue"
end if
RS1.Close
strBgClrSum = "black"

'목표수량을 0으로 초기화
tsum1 = 0
tsum2 = 0
tsum3 = 0
tsum4 = 0
tsum5 = 0
tsum6 = 0
tsum7 = 0

'DB에서 오늘 실적 조회, 파트넘버별로 수량을 조회
SQL = "select PRD_Line, PRD_PartNo, cntPRD_Code = count(PRD_Code), BS_ST = isnull((select BS_ST from tbBOM_Sub where BS_D_No = PRD_PartNo),8) from tbPWS_Raw_Data "
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PRD_BOX_Date = '"&s_Work_Date&"' and "&vbcrlf
SQL = SQL & "	PRD_BOX_Date is not null "&vbcrlf
SQL = SQL & "group by PRD_Line, PRD_PartNo "&vbcrlf
SQL = SQL & "union "&vbcrlf
SQL = SQL & "select PSP_Line, BOM_Sub_BS_D_No, sum(PSP_Count), max(PSP_ST) from tbProcess_State_Plan "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PSP_Sub_YN = 'Y' and len(PSP_Sub_Start) = 4 and len(PSP_Sub_End) = 4 and "&vbcrlf
SQL = SQL & "	PSP_Work_Date = '"&s_Work_Date&"' "&vbcrlf
SQL = SQL & "group by PSP_Line, BOM_Sub_BS_D_No "&vbcrlf

RS1.Open SQL,sys_DBCon

'루핑. 이 과정이 끝나면, 현재 실적수량을 달성하는데 필요한 시간을 차감하게 됨.
do until RS1.eof
	'해당 파트넘버의 실적 수량 및 TT을 변수에 저장
	cntPRD_Code = RS1("cntPRD_Code")	
	BS_ST = RS1("BS_ST")
	if BS_ST = 0 then
		BS_ST = 10
	end if
	if s_Process = "PCBA" then
		'해당하는 라인을 찾아
		if ucase(RS1("PRD_Line")) = "PCBA1" then
			if calcNow1 > 0 then
				if round(calcNow1 / int(BS_ST),0) < int(cntPRD_Code) then	'잔여 누적시간에서 TT으로 할수있는 수량보다, 해당 실적수량이 크다면
					tsum1 = tsum1 + round(calcNow1 / int(BS_ST),0)				'잔여 누적시간에서 TT으로 할수있는 수량만큼을 목표수량에 합산.
				else														'잔여 누적시간에서 TT로 나눈 수량보다 해당 실적수량이 작다면,
					tsum1 = tsum1 + int(cntPRD_Code)							'목표수량에 해당 실적수량만큼 합산.
				end if
				calcNow1 = calcNow1 - (int(cntPRD_Code) * int(BS_ST))	'잔여 누적시간에서 실적수량*TT을 차감
			end if
		elseif ucase(RS1("PRD_Line")) = "PCBA2" then
			if calcNow2 > 0 then
				if round(calcNow2 / int(BS_ST),0) < int(cntPRD_Code) then
					tsum2 = tsum2 + round(calcNow2 / int(BS_ST),0)				
				else														
					tsum2 = tsum2 + int(cntPRD_Code)							
				end if
				calcNow2 = calcNow2 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "PCBA3" then
			if calcNow3 > 0 then
				if round(calcNow3 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum3 = tsum3 + round(calcNow3 / int(BS_ST),0)				
				else														
					tsum3 = tsum3 + int(cntPRD_Code)							
				end if
				calcNow3 = calcNow3 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "PCBA4" then
			if calcNow4 > 0 then
				if round(calcNow4 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum4 = tsum4 + round(calcNow4 / int(BS_ST),0)				
				else														
					tsum4 = tsum4 + int(cntPRD_Code)							
				end if
				calcNow4 = calcNow4 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "PCBA5" then
			if calcNow5 > 0 then
				if round(calcNow5 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum5 = tsum5 + round(calcNow5 / int(BS_ST),0)				
				else														
					tsum5 = tsum5 + int(cntPRD_Code)							
				end if
				calcNow5 = calcNow5 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "PCBA6" then
			if calcNow6 > 0 then
				if round(calcNow6 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum6 = tsum6 + round(calcNow6 / int(BS_ST),0)				
				else														
					tsum6 = tsum6 + int(cntPRD_Code)							
				end if
				calcNow6 = calcNow6 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "PCBA7" then
			if calcNow7 > 0 then
				if round(calcNow7 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum7 = tsum7 + round(calcNow7 / int(BS_ST),0)				
				else														
					tsum7 = tsum7 + int(cntPRD_Code)							
				end if
				calcNow7 = calcNow7 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		end if
	elseif s_Process = "CBOX" then
		if ucase(RS1("PRD_Line")) = "CBOX1" then
			if calcNow1 > 0 then
				if round(calcNow1 / int(BS_ST),0) < int(cntPRD_Code) then	'잔여 누적시간에서 TT으로 할수있는 수량보다, 해당 실적수량이 크다면
					tsum1 = tsum1 + round(calcNow1 / int(BS_ST),0)				'잔여 누적시간에서 TT으로 할수있는 수량만큼을 목표수량에 합산.
				else														'잔여 누적시간에서 TT로 나눈 수량보다 해당 실적수량이 작다면,
					tsum1 = tsum1 + int(cntPRD_Code)							'목표수량에 해당 실적수량만큼 합산.
				end if
				calcNow1 = calcNow1 - (int(cntPRD_Code) * int(BS_ST))	'잔여 누적시간에서 실적수량*TT을 차감
			end if
		elseif ucase(RS1("PRD_Line")) = "CBOX2" then
			if calcNow2 > 0 then
				if round(calcNow2 / int(BS_ST),0) < int(cntPRD_Code) then
					tsum2 = tsum2 + round(calcNow2 / int(BS_ST),0)				
				else														
					tsum2 = tsum2 + int(cntPRD_Code)							
				end if
				calcNow2 = calcNow2 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "CBOX3" then
			if calcNow3 > 0 then
				if round(calcNow3 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum3 = tsum3 + round(calcNow3 / int(BS_ST),0)				
				else														
					tsum3 = tsum3 + int(cntPRD_Code)							
				end if
				calcNow3 = calcNow3 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "CBOX4" then
			if calcNow4 > 0 then
				if round(calcNow4 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum4 = tsum4 + round(calcNow4 / int(BS_ST),0)				
				else														
					tsum4 = tsum4 + int(cntPRD_Code)							
				end if
				calcNow4 = calcNow4 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "CBOX5" then
			if calcNow5 > 0 then
				if round(calcNow5 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum5 = tsum5 + round(calcNow5 / int(BS_ST),0)				
				else														
					tsum5 = tsum5 + int(cntPRD_Code)							
				end if
				calcNow5 = calcNow5 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "CBOX6" then
			if calcNow6 > 0 then
				if round(calcNow6 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum6 = tsum6 + round(calcNow6 / int(BS_ST),0)				
				else														
					tsum6 = tsum6 + int(cntPRD_Code)							
				end if
				calcNow6 = calcNow6 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		elseif ucase(RS1("PRD_Line")) = "CBOX7" then
			if calcNow7 > 0 then
				if round(calcNow7 / int(BS_ST),0) < int(cntPRD_Code) then	
					tsum7 = tsum7 + round(calcNow7 / int(BS_ST),0)				
				else														
					tsum7 = tsum7 + int(cntPRD_Code)							
				end if
				calcNow7 = calcNow7 - (int(cntPRD_Code) * int(BS_ST))	
			end if
		end if
	end if
	RS1.MoveNext
loop
RS1.Close

'남은 시간동안, 남은 계획을 순서대로 수행했을 때, 얼마나 하고 있는 것이 정상인지 파악.
'if s_Process = "PCBA" then
'	if calcNow1 > 0 then
'		tsum1 = tsum1 + getTargetRemainQty("PCBA1", calcNow1)
'	end if
'	if calcNow2 > 0 then
'		tsum2 = tsum2 + getTargetRemainQty("PCBA2", calcNow2)
'	end if
'	if calcNow3 > 0 then
'		tsum3 = tsum3 + getTargetRemainQty("PCBA3", calcNow3)
'	end if
'	if calcNow4 > 0 then
'		tsum4 = tsum4 + getTargetRemainQty("PCBA4", calcNow4)
'	end if
'	if calcNow5 > 0 then
'		tsum5 = tsum5 + getTargetRemainQty("PCBA5", calcNow5)
'	end if
'elseif s_Process = "CBOX" then
'	if calcNow1 > 0 then
'		tsum1 = tsum1 + getTargetRemainQty("CBOX1", calcNow1)
'	end if
'	if calcNow2 > 0 then
'		tsum2 = tsum2 + getTargetRemainQty("CBOX2", calcNow2)
'	end if
'	if calcNow3 > 0 then
'		tsum3 = tsum3 + getTargetRemainQty("CBOX3", calcNow3)
'	end if
'	if calcNow4 > 0 then
'		tsum4 = tsum4 + getTargetRemainQty("CBOX4", calcNow4)
'	end if
'	if calcNow5 > 0 then
'		tsum5 = tsum5 + getTargetRemainQty("CBOX5", calcNow5)
'	end if
'end if

'계획 수량 파악
psum1 = 0
psum2 = 0
psum3 = 0
psum4 = 0
psum5 = 0
psum6 = 0
psum7 = 0
SQL = "select PSP_Line, sumPSP_Count = isnull(sum(PSP_Count),0) from tbProcess_State_Plan where PSP_Work_Date = '"&s_Work_Date&"' group by PSP_Line"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	if s_Process = "PCBA" then
		if RS1("PSP_Line") = "PCBA1" then
			psum1 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "PCBA2" then
			psum2 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "PCBA3" then
			psum3 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "PCBA4" then
			psum4 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "PCBA5" then
			psum5 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "PCBA6" then
			psum6 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "PCBA7" then
			psum7 = RS1("sumPSP_Count")
		end if		
	elseif s_Process = "CBOX" then
		if RS1("PSP_Line") = "CBOX1" then
			psum1 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "CBOX2" then
			psum2 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "CBOX3" then
			psum3 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "CBOX4" then
			psum4 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "CBOX5" then
			psum5 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "CBOX6" then
			psum6 = RS1("sumPSP_Count")
		elseif RS1("PSP_Line") = "CBOX7" then
			psum7 = RS1("sumPSP_Count")
		end if
	end if			

	RS1.MoveNext
loop
RS1.Close

if tsum1 > psum1 then
	tsum1 = psum1
end if
if tsum2 > psum2 then
	tsum2 = psum2
end if
if tsum3 > psum3 then
	tsum3 = psum3
end if
if tsum4 > psum4 then
	tsum4 = psum4
end if
if tsum5 > psum5 then
	tsum5 = psum5
end if
if tsum6 > psum6 then
	tsum6 = psum6
end if
if tsum7 > psum7 then
	tsum7 = psum7
end if

'달성률 초기화
rate1	= "-"
rate2	= "-"
rate3	= "-"
rate4	= "-"
rate5	= "-"
rate6	= "-"
rate7	= "-"
rateSum	= "-"

'각 라인의 목표수량이 0이상이고, 각 실적이 10건 이상일 때만. 달성률 계산
'if tsum1 > 0 and sum1 > 10 then
'	rate1	= round(sum1*100/tsum1,0)
'end if
'if tsum2 > 0 and sum2 > 10 then
'	rate2	= round(sum2*100/tsum2,0)
'end if
'if tsum3 > 0 and sum3 > 10 then
'	rate3	= round(sum3*100/tsum3,0)
'end if
'if tsum4 > 0 and sum4 > 10 then
'	rate4	= round(sum4*100/tsum4,0)
'end if
'if tsum5 > 0 and sum5 > 10 then
'	rate5	= round(sum5*100/tsum5,0)
'end if
'if (tsum1+tsum2+tsum3+tsum4+tsum5) > 0 then
'	rateSum	= round((sum1+sum2+sum3+sum4+sum5)*100/(tsum1+tsum2+tsum3+tsum4+tsum5),0)
'end if

'각 라인의 목표수량이 0이상이고, 각 실적이 10건 이상일 때만. 달성률 계산
if psum1 > 0 and sum1 > 10 then
	rate1	= round(sum1*100/psum1,0)
end if
if psum2 > 0 and sum2 > 10 then
	rate2	= round(sum2*100/psum2,0)
end if
if psum3 > 0 and sum3 > 10 then
	rate3	= round(sum3*100/psum3,0)
end if
if psum4 > 0 and sum4 > 10 then
	rate4	= round(sum4*100/psum4,0)
end if
if psum5 > 0 and sum5 > 10 then
	rate5	= round(sum5*100/psum5,0)
end if
if psum6 > 0 and sum6 > 10 then
	rate6	= round(sum6*100/psum6,0)
end if
if psum7 > 0 and sum7 > 10 then
	rate7	= round(sum7*100/psum7,0)
end if
if (psum1+psum2+psum3+psum4+psum5+psum6+psum7) > 0 then
	rateSum	= round((sum1+sum2+sum3+sum4+sum5+sum6+sum7)*100/(psum1+psum2+psum3+psum4+psum5+psum6+psum7),0)
end if


if isnumeric(rate1) then
	if int(rate1) > 100 then
		rate1 = 100
	end if
end if
if isnumeric(rate2) then
	if int(rate2) > 100 then
		rate2 = 100
	end if
end if
if isnumeric(rate3) then
	if int(rate3) > 100 then
		rate3 = 100
	end if
end if
if isnumeric(rate4) then
	if int(rate4) > 100 then
		rate4 = 100
	end if
end if
if isnumeric(rate5) then
	if int(rate5) > 100 then
		rate5 = 100
	end if
end if
if isnumeric(rate6) then
	if int(rate6) > 100 then
		rate6 = 100
	end if
end if
if isnumeric(rate7) then
	if int(rate7) > 100 then
		rate7 = 100
	end if
end if
set RS1 = nothing
set RS2 = nothing


'70% 이하일 때 적색 표시를 위한 처리
strTRBgClr1 = "black"
if isnumeric(rate1) then
	if int(rate1) < 80 then
		strTRBgClr1 = "darkred"
	end if
	rate1 = rate1 & "%"
end if
strTRBgClr2 = "black"
if isnumeric(rate2) then
	if int(rate2) < 80 then
		strTRBgClr2 = "darkred"
	end if
	rate2 = rate2 & "%"
end if
strTRBgClr3 = "black"
if isnumeric(rate3) then
	if int(rate3) < 80 then
		strTRBgClr3 = "darkred"
	end if
	rate3 = rate3 & "%"
end if
strTRBgClr4 = "black"
if isnumeric(rate4) then
	if int(rate4) < 80 then
		strTRBgClr4 = "darkred"
	end if
	rate4 = rate4 & "%"
end if
strTRBgClr5 = "black"
if isnumeric(rate5) then
	if int(rate5) < 80 then
		strTRBgClr5 = "darkred"
	end if
	rate5 = rate5 & "%"
end if
strTRBgClr6 = "black"
if isnumeric(rate6) then
	if int(rate6) < 80 then
		strTRBgClr6 = "darkred"
	end if
	rate6 = rate6 & "%"
end if
strTRBgClr7 = "black"
if isnumeric(rate7) then
	if int(rate7) < 80 then
		strTRBgClr7 = "darkred"
	end if
	rate7 = rate7 & "%"
end if
strTRBgClrSum = "black"
if isnumeric(rateSum) then
	if int(rateSum) < 80 then
		strTRBgClrSum = "darkred"
	end if
	rateSum = rateSum & "%"
end if
%>

<%
'계획과 실적을 비교해서, 실적을 제외한 계획을 기반으로 남은시간에 할 수 있는 작업량을 구하는 함수.
function getTargetRemainQty(strLine, calcNow)
	dim arrPWS_Raw_Data
	
	dim strBOM_Sub_BS_D_No
	dim strPSP_Count
	dim strPSP_ST
	dim arrBOM_Sub_BS_D_No
	dim arrPSP_Count
	dim arrPSP_ST
	
	dim TargetRemainQty
	
	dim CNT1
	dim CNT2
	
	dim RS1
	
	'해당라인의 실적 정보를 가져옴
	arrPWS_Raw_Data		= getPWS_Raw_Data(s_Work_Date, strLine)	
	
	'해당라인의 계획 목록을 루핑
	SQL = "select BOM_Sub_BS_D_No, PSP_Count, PSP_ST "&vbcrlf
	SQL = SQL & "from tbProcess_State_Plan "&vbcrlf
	SQL = SQL & "where "&vbcrlf 
	SQL = SQL & "	PSP_Work_Date = '"&s_Work_Date&"' and "&vbcrlf
	SQL = SQL & "	PSP_Line = '"&strLine&"' and "&vbcrlf
	SQL = SQL & "	(PSP_Sub_YN <> 'Y' or PSP_Sub_YN is null or len(PSP_Sub_Start) <> 4 or PSP_Sub_Start is null or len(PSP_Sub_End) <> 4 or PSP_Sub_End is null) "&vbcrlf
	SQL = SQL & "order by PSP_Code "&vbcrlf
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		'계획 PNO와, 계획수량, TT를 배열화 함.
		strBOM_Sub_BS_D_No	= strBOM_Sub_BS_D_No	& RS1("BOM_Sub_BS_D_No")	&"|"
		strPSP_Count		= strPSP_Count			& RS1("PSP_Count")			&"|"
		strPSP_ST			= strPSP_ST				& RS1("PSP_ST")				&"|"
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
	'다음배열은 오늘 계획의 PNO, 수량, TT정보들이 담김
	arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,"|")
	arrPSP_Count		= split(strPSP_Count,"|")
	arrPSP_ST			= split(strPSP_ST,"|")

	'계획 배열을 루핑 하면서, 실적 배열과 교차 비교. 한마디로 계획에 일치하는 실적이 있는 경우. 계획에서 제거하기 위함.
	for CNT1=0 to ubound(arrBOM_Sub_BS_D_No) - 1
		for CNT2=0 to ubound(arrPWS_Raw_Data)
			'계획과 실적의 파트넘버가 동일하고, 계획과 실적 수량이 0 이상일 경우.
			if arrPWS_Raw_Data(CNT2,0) = arrBOM_Sub_BS_D_No(CNT1) and arrPWS_Raw_Data(CNT2,1) > 0 and arrPSP_Count(CNT1) > 0 then
				if arrPWS_Raw_Data(CNT2,1) = arrPSP_Count(CNT1) then '수량이 서로 같다면, 계획수량과 실적수량 0으로 변경
					arrPWS_Raw_Data(CNT2,1) = 0
					arrPSP_Count(CNT1) = 0
				elseif arrPWS_Raw_Data(CNT2,1) > arrPSP_Count(CNT1) then '실적수량이 더 크다면, 실적에는 잔량만 남기고, 계획수량은 0으로 수정
					arrPWS_Raw_Data(CNT2,1) = arrPWS_Raw_Data(CNT2,1) - arrPSP_Count(CNT1)
					arrPSP_Count(CNT1) = 0
				elseif arrPWS_Raw_Data(CNT2,1) < arrPSP_Count(CNT1) then '계획수량이 더 크다면, 계획수량에서 실적수량을 차감하고, 실적수량은 0으로 소진.
					arrPSP_Count(CNT1) = arrPSP_Count(CNT1) - arrPWS_Raw_Data(CNT2,1)
					arrPWS_Raw_Data(CNT2,1) = 0
				end if
			end if
		next
	next
	
	'실적을 차감한 계획배열을 루핑
	for CNT1=0 to ubound(arrBOM_Sub_BS_D_No) - 1
		'해당계획의 계획수량만큼 루핑
		for CNT2=0 to arrPSP_Count(CNT1)
			if calcNow > 0 then '잔여시간이 남아있다면,
				calcNow = calcNow - arrPSP_ST(CNT1) '잔여시간에서 TT를 차감
				TargetRemainQty = TargetRemainQty + 1 '잔여시간을 1증가 시킴
			end if
		next
	next
	
	getTargetRemainQty = TargetRemainQty
end function
%>

<script language="javascript">
<%
dim strLineTitle

if s_Process ="PCBA" then
	strLineTitle = "P"
elseif s_Process = "CBOX" then
	strLineTitle = "C"
end if
%>
var nSum1 = "<%=sum1%>";
var nSum2 = "<%=sum2%>";
var nSum3 = "<%=sum3%>";
var nSum4 = "<%=sum4%>";
var nSum5 = "<%=sum5%>";
var nSumTotal = "<%=sum1+sum2+sum3+sum4+sum5%>";

var strHTML = "";
strHTML += "<table width=100% cellpadding=0 cellspacing=1 bgcolor='white' style='color:white;font-size:80px;text-align:center;font-weight:bold'>";
strHTML += "<col width=200px></col>";
strHTML += "<col width=250px></col>";
strHTML += "<col width=250px></col>";
strHTML += "<col width=250px></col>";
strHTML += "<col width=250px></col>";
strHTML += "<col></col>";
strHTML += "<tr bgcolor=<%=strTRBgClr1%>>";
strHTML += "	<td><%=strLineTitle%>M1</td>";
strHTML += "	<td align=right><%=psum1%></td>";
strHTML += "	<td align=right><%=isum1%></td>";
strHTML += "	<td align=right>"+nSum1+"</td>";
strHTML += "	<td align=right><%=rate1%></td>";
strHTML += "	<td bgcolor=<%=strBgClr1%>><%=strLineState1%></td>";
strHTML += "</tr>";
strHTML += "<tr bgcolor=<%=strTRBgClr2%>>";
strHTML += "	<td><%=strLineTitle%>M2</td>";
strHTML += "	<td align=right><%=psum2%></td>";
strHTML += "	<td align=right><%=isum2%></td>";
strHTML += "	<td align=right>"+nSum2+"</td>";
strHTML += "	<td align=right><%=rate2%></td>";
strHTML += "	<td bgcolor=<%=strBgClr2%>><%=strLineState2%></td>";
strHTML += "</tr>";
strHTML += "<tr bgcolor=<%=strTRBgClr3%>>";
strHTML += "	<td><%=strLineTitle%>M3</td>";
strHTML += "	<td align=right><%=psum3%></td>";
strHTML += "	<td align=right><%=isum3%></td>";
strHTML += "	<td align=right>"+nSum3+"</td>";
strHTML += "	<td align=right><%=rate3%></td>";
strHTML += "	<td bgcolor=<%=strBgClr3%>><%=strLineState3%></td>";
strHTML += "</tr>";
strHTML += "<tr bgcolor=<%=strTRBgClr4%>>";
strHTML += "	<td><%=strLineTitle%>M4</td>";
strHTML += "	<td align=right><%=psum4%></td>";
strHTML += "	<td align=right><%=isum4%></td>";
strHTML += "	<td align=right>"+nSum4+"</td>";
strHTML += "	<td align=right><%=rate4%></td>";
strHTML += "	<td bgcolor=<%=strBgClr4%>><%=strLineState4%></td>";
strHTML += "</tr>";
strHTML += "<tr bgcolor=<%=strTRBgClr5%>>";
strHTML += "	<td><%=strLineTitle%>M5</td>";
strHTML += "	<td align=right><%=psum5%></td>";
strHTML += "	<td align=right><%=isum5%></td>";
strHTML += "	<td align=right>"+nSum5+"</td>";
strHTML += "	<td align=right><%=rate5%></td>";
strHTML += "	<td bgcolor=<%=strBgClr5%>><%=strLineState5%></td>";
strHTML += "</tr>";
strHTML += "<tr bgcolor=<%=strTRBgClrSum%>>";
strHTML += "	<td>SUM</td>";
strHTML += "	<td align=right><%=psum1+psum2+psum3+psum4+psum5%></td>";
strHTML += "	<td align=right><%=isum1+isum2+isum3+isum4+isum5%></td>";
strHTML += "	<td align=right>"+nSumTotal+"</td>";
strHTML += "	<td align=right><%=rateSum%></td>";
strHTML += "	<td bgcolor=<%=strBgClrSum%>>&nbsp;</td>";
strHTML += "</tr>";
strHTML += "<tr bgcolor=black>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "</tr>";
strHTML += "</table>";

parent.idContent.innerHTML = strHTML;

function reload_handle()
{
<%
	if Request("s_Multi_YN") <> "Y" then
%>
		location.reload();
<%
	elseif s_Process = "CBOX" then
%>
		location.href='new_mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=PCBA&s_Work_Date=<%=s_Work_Date%>'
<%
	else
%>
		location.href='new_mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=CBOX&s_Work_Date=<%=s_Work_Date%>'
<%
	end if
%>
}

/*
<%
if Request("s_Multi_YN") = "Y" then
%>
function fRun()
{
	if(document.readyState == "complete")
	{
<%
	if s_Process = "CBOX" then
%>
		location.href='new_mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=PCBA&s_Work_Date=<%=s_Work_Date%>'
<%
	else
%>
		location.href='new_mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=CBOX&s_Work_Date=<%=s_Work_Date%>'
<%
	end if
%>
	}
	else
	{
		setTimeout("fRun()",30000);
	}
}
fRun();
<%
else
%>
function fRun()
{
	if(document.readyState == "complete")
	{
		location.reload();
	}
	else
	{
		setTimeout("fRun()",30000);
	}
}
fRun();
<%
end if
%>
*/
</script>


<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	
	