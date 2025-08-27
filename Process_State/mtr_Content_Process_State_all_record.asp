<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
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

'생산시작시간 
dim calcPRD_Start1
dim calcPRD_Start2
dim calcPRD_Start3
dim calcPRD_Start4
dim calcPRD_Start5
dim calcPRD_Start6
dim calcPRD_Start7

'목표 변수 
dim tsum1
dim tsum2
dim tsum3
dim tsum4
dim tsum5
dim tsum6
dim tsum7


'무작업 변수 
dim idlesum1
dim idlesum2
dim idlesum3
dim idlesum4
dim idlesum5
dim idlesum6
dim idlesum7
dim idlesumTotal
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

dim strSimilar
dim arrSimilar
dim arrSimilarDetail
strSimilar = strSimilar & "EBR715081$-EBR741529-//"
strSimilar = strSimilar & "EBR644383$-EBR662511-EBR737344-//"
strSimilar = strSimilar & "EBR624063$-EBR654006-//"
strSimilar = strSimilar & "EBR391877$-EBR622045-EBR784017-EBR806189-EBR813330-//"
strSimilar = strSimilar & "6871A10233$-EBR535783-EBR610631-//"
strSimilar = strSimilar & "6871A20181$-EBR515892-EBR515959-//"
strSimilar = strSimilar & "EBR337500$-EBR420488-EBR489280-EBR568373-EBR622537-//"
strSimilar = strSimilar & "EBR351584$-EBR412888-EBR420085-EBR442026-EBR564637-EBR577053-EBR577370-EBR618909-EBR740459-EBR743984-EBR775510-EBR775954-EBR775955-EBR779234-EBR779637-EBR784020-EBR788952-EBR792664-EBR792665-EBR798438-EBR801090-EBR815512-//"
strSimilar = strSimilar & "EBR355200$-EBR399048-EBR631040-EBR715171-EBR774722-EBR794405-//"
strSimilar = strSimilar & "6871A20156$-EBR356394-EBR441696-EBR604434-//"
strSimilar = strSimilar & "6871A20889$-6871A20891-//"
strSimilar = strSimilar & "6871A10161$-6871A20684-//"
strSimilar = strSimilar & "6871A10117$-6871A20679-EBR431272-//"
strSimilar = strSimilar & "6871A20272$-6871A20547-//"
strSimilar = strSimilar & "6871A20294$-6871A20309-6871A20310-6871A20311-6871A20312-6871A20373-6871A20493-6871A20494-6871A20495-6871A20562-6871A20565-//"
strSimilar = strSimilar & "6871A20225$-6871A20229-6871A20235-//"
strSimilar = strSimilar & "6871A20107$-6871A20222-//"
strSimilar = strSimilar & "6871A20164$-6871A20216-6871A20218-6871A20220-6871A20240-//"
strSimilar = strSimilar & "6871A20146$-6871A20160-6871A20212-6871A20232-6871A20352-//"
strSimilar = strSimilar & "6871A10042$-6871A20040-6871A20067-6871A20082-6871A20152-6871A20158-6871A20415-//"
strSimilar = strSimilar & "6871A20007$-6871A20008-//"
strSimilar = strSimilar & "6871A10231$-6871A10362-6871A10363-6871A10366-EBR341635-//"
strSimilar = strSimilar & "6871A10158$-6871A10209-6871A10338-6871A10370-//"
strSimilar = strSimilar & "6871A10108$-6871A10167-//"
strSimilar = strSimilar & "6871A10105$-6871A10165-//"
strSimilar = strSimilar & "6871A10056$-6871A10143-//"
strSimilar = strSimilar & "6871A00089$-6871A10140-6871A10342-//"
strSimilar = strSimilar & "6871A01002$-6871A10070-6871A10188-6871A20188-EBR615952-//"
strSimilar = strSimilar & "6871A10008$-6871A10038-6871A10040-6871A10109-6871A10116-//"
strSimilar = strSimilar & "6871A10009$-6871A10020-6871A10023-6871A10026-6871A10030-6871A10048-//"
strSimilar = strSimilar & "6871A00012$-6871A00090-6871A10131-//"
strSimilar = strSimilar & "6871A00007$-6871A00009-6871A10089-6871A10106-6871A10107-6871A10124-6871A10125-6871A10148-6871A10166-6871A10187-6871A10217-//"
arrSimilar = split(strSimilar,"//")

'SQL = "insert into tbTest_setinterval (ts_Work,ts_Desc,ts_Now,ts_Diff) values ('ProcessState','ALL',getdate(),0)"
'sys_DBCon.execute(SQL)

'날짜는 무조건 현재 날짜로
's_Work_Date = request("s_Work_Date")
's_Work_Date = "2015-08-17"
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

'라인명이 불분명한 기록을 조회
SQL =		"select PRD_Code, PRD_PartNo from tbPWS_Raw_Data "
SQL = SQL & "where PRD_Line not in ('pcba1','pcba2','pcba3','pcba4','pcba5','pcba6','pcba7','cbox1','cbox2','cbox3','cbox4','cbox5','cbox6','cbox7') and "
SQL = SQL & "(PRD_ICT_Date = '"&s_Work_Date&"' or PRD_BOX_Date = '"&s_Work_Date&"') and PRD_Dummy_YN is null "&vbcrlf
'response.write SQL &"<BR>"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else
	dim strPSP
	strPSP = "-"
	SQL = "select PSP_Line, BOM_Sub_BS_D_No from tbProcess_State_Plan where PSP_Work_Date = '"&s_Work_Date&"'"	
	RS2.Open SQL,sys_DBCon
	Do until RS2.Eof
		If InStr(strPSP,"-"&RS2(0)&RS2(1)&"-") = 0 Then
			strPSP = strPSP &RS2(0)&RS2(1)&"-"
		Else
			strPSP = replace(strPSP,RS2(0)&RS2(1),"-")
		End If
		RS2.MoveNext
	loop
	RS2.Close
	do until RS1.Eof
		If InStr(strPSP,RS1("PRD_PartNo")) > 0 Then
			SQL = "update tbPWS_Raw_Data set PRD_Line = '"&mid(strPSP,InStr(strPSP,RS1("PRD_PartNo"))-5,5)&"' where PRD_Code = "&RS1("PRD_Code")
			sys_DBCon.execute(SQL)
		End If
		RS1.MoveNext
	loop
end if
RS1.Close	
		
		



'라인별로, 포장 완료 수량을 집계 (투입 후킹 기준)
'if s_Process = "CBOX" then 
	SQL = "select PRD_Line, cntPRD_Code = count(PRD_Code) from tbPWS_Raw_Data "
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	(PRD_ICT_Date = '"&s_Work_Date&"' or PRD_BOX_Date = '"&s_Work_Date&"') and PRD_Dummy_YN is null "&vbcrlf
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
'end if

'임시 실적(계획으로 라인 찾는 방식)
if s_Process = "temp" then 
	dim CNT1
	dim strPSP_Line
	dim strBOM_Sub_BS_D_No
	dim arrPSP_Line
	dim arrBOM_Sub_BS_D_No
	SQL = 		"select PSP_Line, BOM_Sub_BS_D_No from tbProcess_State_Plan "&vbcrlf
	SQL = SQL & " where PSP_Work_Date = '"&s_Work_Date&"' and "&vbcrlf
	SQL = SQL & " 	PSP_Line in ('PCBA1','PCBA2','PCBA3','PCBA4','PCBA5','PCBA6','PCBA7') "&vbcrlf	
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		strPSP_Line = strPSP_Line & RS1("PSP_Line") &"||"
		strBOM_Sub_BS_D_No = strBOM_Sub_BS_D_No & RS1("BOM_Sub_BS_D_No") &"||"
		RS1.MoveNext
	loop
	RS1.Close
	arrPSP_Line			= split(strPSP_Line,"||")
	arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,"||")
	
	sum1 = 0
	sum2 = 0
	sum3 = 0
	sum4 = 0
	sum5 = 0
	sum6 = 0
	sum7 = 0
	SQL = 		"select PRD_PartNo, cntPRD_Code = count(PRD_Code) from tbPWS_Raw_Data "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	if s_Process = "PCBA" then 
		SQL = SQL & "	PRD_ICT_Date = '"&s_Work_Date&"' and "&vbcrlf
		SQL = SQL & "	PRD_ICT_Date <> '' and PRD_ICT_Date is not null and PRD_Dummy_YN is null "&vbcrlf
	else
		SQL = SQL & "	PRD_BOX_Date = '"&s_Work_Date&"' and "&vbcrlf
		SQL = SQL & "	PRD_BOX_Date <> '' and PRD_BOX_Date is not null and PRD_Dummy_YN is null "&vbcrlf
	end if
	SQL = SQL & "group by PRD_PartNo"&vbcrlf
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		for CNT1 = 0 to ubound(arrPSP_Line)-1
			if arrBOM_Sub_BS_D_No(CNT1) = RS1("PRD_PartNo") then
				if arrPSP_Line(CNT1) = "PCBA1" then
					sum1 = sum1 + RS1("cntPRD_Code")
					exit for
				elseif arrPSP_Line(CNT1) = "PCBA2" then
					sum2 = sum2 + RS1("cntPRD_Code")
					exit for
				elseif arrPSP_Line(CNT1) = "PCBA3" then
					sum3 = sum3 + RS1("cntPRD_Code")
					exit for
				elseif arrPSP_Line(CNT1) = "PCBA4" then
					sum4 = sum4 + RS1("cntPRD_Code")
					exit for
				elseif arrPSP_Line(CNT1) = "PCBA5" then
					sum5 = sum5 + RS1("cntPRD_Code")
					exit for
				elseif arrPSP_Line(CNT1) = "PCBA6" then
					sum6 = sum6 + RS1("cntPRD_Code")
					exit for
				elseif arrPSP_Line(CNT1) = "PCBA7" then
					sum7 = sum7 + RS1("cntPRD_Code")
					exit for
				end if
			end if
		next
		RS1.MoveNext
	loop
	RS1.Close
	SQL = ""
end if

'투입때문에 추가
'SQL = "select PRD_Line, cntPRD_Code = count(PRD_Code) from tbPWS_Raw_Data "
'SQL = SQL & "where "&vbcrlf
'SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and PRD_Dummy_YN is null "&vbcrlf
'SQL = SQL & "group by PRD_Line"&vbcrlf
'RS1.Open SQL,sys_DBCon
'do until RS1.Eof
'	if s_Process = "PCBA" then 
'		if ucase(RS1("PRD_Line")) = "PCBA1" then
'			isum1 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "PCBA2" then
'			isum2 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "PCBA3" then
'			isum3 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "PCBA4" then
'			isum4 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "PCBA5" then
'			isum5 = RS1("cntPRD_Code")
'		end if					
'	elseif s_Process = "CBOX" then 
'		if ucase(RS1("PRD_Line")) = "CBOX1" then
'			isum1 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "CBOX2" then
'			isum2 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "CBOX3" then
'			isum3 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "CBOX4" then
'			isum4 = RS1("cntPRD_Code")
'		elseif ucase(RS1("PRD_Line")) = "CBOX5" then
'			isum5 = RS1("cntPRD_Code")
'		end if	
'	end if
'	RS1.MoveNext
'loop
'RS1.Close

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
calcNow = calcNow * 60
calcNow = getRestedCalcNow(calcNow)



'초로 바꿈


'각 라인별로 사용하기 위해서 변수별 분배 및 초로 환산
calcNow1 = calcNow
calcNow2 = calcNow
calcNow3 = calcNow
calcNow4 = calcNow
calcNow5 = calcNow
calcNow6 = calcNow
calcNow7 = calcNow

'각 라인별 최초 생산 시간으로 구하기
SQL = "select PRD_Line, minPRD_Input_Time = min(PRD_Input_Time) from tbPWS_Raw_Data "
SQL = SQL & "where "&vbcrlf
if s_Process = "PCBA" then 
	SQL = SQL & "	PRD_LINE like 'pcba%' and "&vbcrlf
else
	SQL = SQL & "	PRD_LINE like 'cbox%' and "&vbcrlf
end if
SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and "&vbcrlf
SQL = SQL & "	PRD_Input_Date is not null"&vbcrlf
SQL = SQL & "group by PRD_Line"&vbcrlf
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	if right(RS1("PRD_Line"),1) = "1" then
		calcPRD_Start1 = getRestedCalcNow(getCalcPRD_Start(RS1("minPRD_Input_Time")))
	elseif right(RS1("PRD_Line"),1) = "2" then
		calcPRD_Start2 = getRestedCalcNow(getCalcPRD_Start(RS1("minPRD_Input_Time")))
	elseif right(RS1("PRD_Line"),1) = "3" then
		calcPRD_Start3 = getRestedCalcNow(getCalcPRD_Start(RS1("minPRD_Input_Time")))
	elseif right(RS1("PRD_Line"),1) = "4" then
		calcPRD_Start4 = getRestedCalcNow(getCalcPRD_Start(RS1("minPRD_Input_Time")))
	elseif right(RS1("PRD_Line"),1) = "5" then
		calcPRD_Start5 = getRestedCalcNow(getCalcPRD_Start(RS1("minPRD_Input_Time")))
	elseif right(RS1("PRD_Line"),1) = "6" then
		calcPRD_Start6 = getRestedCalcNow(getCalcPRD_Start(RS1("minPRD_Input_Time")))
	elseif right(RS1("PRD_Line"),1) = "7" then
		calcPRD_Start7 = getRestedCalcNow(getCalcPRD_Start(RS1("minPRD_Input_Time")))
	end if
					
	RS1.MoveNext
loop
RS1.Close
'response.write calcPRD_Start1&"/"&calcPRD_Start2&"/"&calcPRD_Start3&"/"&calcPRD_Start4&"/"&calcPRD_Start5&"<br>"

function getCalcPRD_Start(minPRD_Input_Time)
	dim calcPRD_Start
	'최초 생산시간을 초로 환산
	calcPRD_Start = (int(left(minPRD_Input_Time,2)*60) + int(right(minPRD_Input_Time,2)))*60
	if calcPRD_Start < 30000 then '8시 20분 이전에 생산이 시작되었다면
		calcPRD_Start = 30000 '그냥 8시 20분으로 보정
	end if
	getCalcPRD_Start = calcPRD_Start
end function

function getRestedCalcNow(calcNow)
	'쉬는 시간 중이라면, 쉬는 시간 시작 상태로 고정
	if calcNow > 620*60 and calcNow <= 630*60 then '10시 20분 ~ 30분 
		calcNow = 620*60
	end if
	if calcNow > 750*60 and calcNow <= 790*60 then '12시 30분~13시 10분 
		calcNow = 750*60
	end if
	if calcNow > 910*60 and calcNow <= 920*60 then '3시 10분 ~ 20분 
		calcNow = 910*60
	end if
	if calcNow > 1040*60 and calcNow <= 1060*60 then '5시 20분~40분
		calcNow = 1040*60
	end if
	
	'쉬는 시간을 거친 수 만큼 쉬는 시간 차감
	if calcNow > 1060*60 then '17시 40분
		calcNow = calcNow - (20+10+40+10)*60
	elseif calcNow > 920*60 then '15시 20분 오전한개, 점심한개, 오후 한개 지남 
		calcNow = calcNow - (10+40+10)*60
	elseif calcNow > 790*60 then '13시 10분 오전쉬는시간 + 점심지남
		calcNow = calcNow - (40+10)*60
	elseif calcNow > 630*60 then '10시 30분 오전쉬는시간 하나 지남
		calcNow = calcNow - 10*60
	end if
	getRestedCalcNow = calcNow
end function

calcNow1 = getCalcNow(calcNow1, calcPRD_Start1)
calcNow2 = getCalcNow(calcNow2, calcPRD_Start2)
calcNow3 = getCalcNow(calcNow3, calcPRD_Start3)
calcNow4 = getCalcNow(calcNow4, calcPRD_Start4)
calcNow5 = getCalcNow(calcNow5, calcPRD_Start5)
calcNow6 = getCalcNow(calcNow6, calcPRD_Start6)
calcNow7 = getCalcNow(calcNow7, calcPRD_Start7)

function getCalcNow(calcNow, getCalcPRD_Start)
	if getCalcPRD_Start = "" then '생산시작을 안했다면 지금이 몇시든 간에 생산가능 시간은 0
		getCalcNow = 0 
	elseif getCalcPRD_Start > calcNow then '이론상 생산시간이 지금시간보다 뒤는 될 수 없다, 하지만 그렇다면, 생산가능 시간은 0
		getCalcNow = 0 
	elseif getCalcPRD_Start < calcNow then '예를 들어 지금 9시 인데, 생산시작이 8시 40분이었다, 그렇다면 생산가능 시간은 20분
		getCalcNow = calcNow - getCalcPRD_Start
	end if
end function

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

if s_Process = "PCBA" then
	tsum1 = getTargetQty("PCBA1",calcNow1)
	tsum2 = getTargetQty("PCBA2",calcNow2)
	tsum3 = getTargetQty("PCBA3",calcNow3)
	tsum4 = getTargetQty("PCBA4",calcNow4)
	tsum5 = getTargetQty("PCBA5",calcNow5)
	tsum6 = getTargetQty("PCBA6",calcNow6)
	tsum7 = getTargetQty("PCBA7",calcNow7)
elseif s_Process = "CBOX" then
	tsum1 = getTargetQty("CBOX1",calcNow1)
	tsum2 = getTargetQty("CBOX2",calcNow2)
	tsum3 = getTargetQty("CBOX3",calcNow3)
	tsum4 = getTargetQty("CBOX4",calcNow4)
	tsum5 = getTargetQty("CBOX5",calcNow5)
	tsum6 = getTargetQty("CBOX6",calcNow6)
	tsum7 = getTargetQty("CBOX7",calcNow7)
end if

'DB에서 오늘 실적 조회, 파트넘버별로 수량을 조회
'SQL = "select PRD_Line, PRD_PartNo, cntPRD_Code = count(PRD_Code), BS_ST = isnull((select BS_ST from tbBOM_Sub where BS_D_No = PRD_PartNo),8) from tbPWS_Raw_Data "
'SQL = SQL & "where "&vbcrlf
'SQL = SQL & "	PRD_BOX_Date = '"&s_Work_Date&"' and "&vbcrlf
'SQL = SQL & "	PRD_BOX_Date is not null "&vbcrlf
'SQL = SQL & "group by PRD_Line, PRD_PartNo "&vbcrlf
'SQL = SQL & "union "&vbcrlf
'SQL = SQL & "select PSP_Line, BOM_Sub_BS_D_No, sum(PSP_Count), max(PSP_ST) from tbProcess_State_Plan "&vbcrlf
'SQL = SQL & "where "&vbcrlf
'SQL = SQL & "	PSP_Sub_YN = 'Y' and len(PSP_Sub_Start) = 4 and len(PSP_Sub_End) = 4 and "&vbcrlf
'SQL = SQL & "	PSP_Work_Date = '"&s_Work_Date&"' "&vbcrlf
'SQL = SQL & "group by PSP_Line, BOM_Sub_BS_D_No "&vbcrlf
'RS1.Open SQL,sys_DBCon

'루핑. 이 과정이 끝나면, 현재 실적수량을 달성하는데 필요한 시간을 차감하게 됨.
'do until RS1.eof
'	'해당 파트넘버의 실적 수량 및 TT을 변수에 저장
'	cntPRD_Code = RS1("cntPRD_Code")	
'	BS_ST = RS1("BS_ST")
'	if BS_ST = 0 then
'		BS_ST = 10
'	end if
'	if s_Process = "PCBA" then
'		'해당하는 라인을 찾아
'		if ucase(RS1("PRD_Line")) = "PCBA1" then
'			if calcNow1 > 0 then
'				if round(calcNow1 / int(BS_ST),0) < int(cntPRD_Code) then	'잔여 누적시간에서 TT으로 할수있는 수량보다, 해당 실적수량이 크다면
'					tsum1 = tsum1 + round(calcNow1 / int(BS_ST),0)				'잔여 누적시간에서 TT으로 할수있는 수량만큼을 목표수량에 합산.
'				else														'잔여 누적시간에서 TT로 나눈 수량보다 해당 실적수량이 작다면,
'					tsum1 = tsum1 + int(cntPRD_Code)							'목표수량에 해당 실적수량만큼 합산.
'				end if
'				calcNow1 = calcNow1 - (int(cntPRD_Code) * int(BS_ST))	'잔여 누적시간에서 실적수량*TT을 차감
'			end if
'		elseif ucase(RS1("PRD_Line")) = "PCBA2" then
'			if calcNow2 > 0 then
'				if round(calcNow2 / int(BS_ST),0) < int(cntPRD_Code) then
'					tsum2 = tsum2 + round(calcNow2 / int(BS_ST),0)				
'				else														
'					tsum2 = tsum2 + int(cntPRD_Code)							
'				end if
'				calcNow2 = calcNow2 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		elseif ucase(RS1("PRD_Line")) = "PCBA3" then
'			if calcNow3 > 0 then
'				if round(calcNow3 / int(BS_ST),0) < int(cntPRD_Code) then	
'					tsum3 = tsum3 + round(calcNow3 / int(BS_ST),0)				
'				else														
'					tsum3 = tsum3 + int(cntPRD_Code)							
'				end if
'				calcNow3 = calcNow3 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		elseif ucase(RS1("PRD_Line")) = "PCBA4" then
'			if calcNow4 > 0 then
'				if round(calcNow4 / int(BS_ST),0) < int(cntPRD_Code) then	
'					tsum4 = tsum4 + round(calcNow4 / int(BS_ST),0)				
'				else														
'					tsum4 = tsum4 + int(cntPRD_Code)							
'				end if
'				calcNow4 = calcNow4 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		elseif ucase(RS1("PRD_Line")) = "PCBA5" then
'			if calcNow5 > 0 then
'				if round(calcNow5 / int(BS_ST),0) < int(cntPRD_Code) then	
'					tsum5 = tsum5 + round(calcNow5 / int(BS_ST),0)				
'				else														
'					tsum5 = tsum5 + int(cntPRD_Code)							
'				end if
'				calcNow5 = calcNow5 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		end if
'	elseif s_Process = "CBOX" then
'		if ucase(RS1("PRD_Line")) = "CBOX1" then
'			if calcNow1 > 0 then
'				if round(calcNow1 / int(BS_ST),0) < int(cntPRD_Code) then	'잔여 누적시간에서 TT으로 할수있는 수량보다, 해당 실적수량이 크다면
'					tsum1 = tsum1 + round(calcNow1 / int(BS_ST),0)				'잔여 누적시간에서 TT으로 할수있는 수량만큼을 목표수량에 합산.
'				else														'잔여 누적시간에서 TT로 나눈 수량보다 해당 실적수량이 작다면,
'					tsum1 = tsum1 + int(cntPRD_Code)							'목표수량에 해당 실적수량만큼 합산.
'				end if
'				calcNow1 = calcNow1 - (int(cntPRD_Code) * int(BS_ST))	'잔여 누적시간에서 실적수량*TT을 차감
'			end if
'		elseif ucase(RS1("PRD_Line")) = "CBOX2" then
'			if calcNow2 > 0 then
'				if round(calcNow2 / int(BS_ST),0) < int(cntPRD_Code) then
'					tsum2 = tsum2 + round(calcNow2 / int(BS_ST),0)				
'				else														
'					tsum2 = tsum2 + int(cntPRD_Code)							
'				end if
'				calcNow2 = calcNow2 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		elseif ucase(RS1("PRD_Line")) = "CBOX3" then
'			if calcNow3 > 0 then
'				if round(calcNow3 / int(BS_ST),0) < int(cntPRD_Code) then	
'					tsum3 = tsum3 + round(calcNow3 / int(BS_ST),0)				
'				else														
'					tsum3 = tsum3 + int(cntPRD_Code)							
'				end if
'				calcNow3 = calcNow3 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		elseif ucase(RS1("PRD_Line")) = "CBOX4" then
'			if calcNow4 > 0 then
'				if round(calcNow4 / int(BS_ST),0) < int(cntPRD_Code) then	
'					tsum4 = tsum4 + round(calcNow4 / int(BS_ST),0)				
'				else														
'					tsum4 = tsum4 + int(cntPRD_Code)							
'				end if
'				calcNow4 = calcNow4 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		elseif ucase(RS1("PRD_Line")) = "CBOX5" then
'			if calcNow5 > 0 then
'				if round(calcNow5 / int(BS_ST),0) < int(cntPRD_Code) then	
'					tsum5 = tsum5 + round(calcNow5 / int(BS_ST),0)				
'				else														
'					tsum5 = tsum5 + int(cntPRD_Code)							
'				end if
'				calcNow5 = calcNow5 - (int(cntPRD_Code) * int(BS_ST))	
'			end if
'		end if
'	end if
'	RS1.MoveNext
'loop
'RS1.Close

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
if tsum1 > 0 and sum1 > 10 then
	rate1	= round(sum1*100/tsum1,0)
end if
if tsum2 > 0 and sum2 > 10 then
	rate2	= round(sum2*100/tsum2,0)
end if
if tsum3 > 0 and sum3 > 10 then
	rate3	= round(sum3*100/tsum3,0)
end if
if tsum4 > 0 and sum4 > 10 then
	rate4	= round(sum4*100/tsum4,0)
end if
if tsum5 > 0 and sum5 > 10 then
	rate5	= round(sum5*100/tsum5,0)
end if
if tsum6 > 0 and sum6 > 10 then
	rate6	= round(sum6*100/tsum6,0)
end if
if tsum7 > 0 and sum7 > 10 then
	rate7	= round(sum7*100/tsum7,0)
end if
if (tsum1+tsum2+tsum3+tsum4+tsum5+tsum6+tsum7) > 0 then
	rateSum	= round((sum1+sum2+sum3+sum4+sum5+sum6+sum7)*100/(tsum1+tsum2+tsum3+tsum4+tsum5+tsum6+tsum7),0)
end if

'각 라인의 목표수량이 0이상이고, 각 실적이 10건 이상일 때만. 달성률 계산
'if psum1 > 0 and sum1 > 10 then
'	rate1	= round(sum1*100/psum1,0)
'end if
'if psum2 > 0 and sum2 > 10 then
'	rate2	= round(sum2*100/psum2,0)
'end if
'if psum3 > 0 and sum3 > 10 then
'	rate3	= round(sum3*100/psum3,0)
'end if
'if psum4 > 0 and sum4 > 10 then
'	rate4	= round(sum4*100/psum4,0)
'end if
'if psum5 > 0 and sum5 > 10 then
'	rate5	= round(sum5*100/psum5,0)
'end if
'if (psum1+psum2+psum3+psum4+psum5) > 0 then
'	rateSum	= round((sum1+sum2+sum3+sum4+sum5)*100/(psum1+psum2+psum3+psum4+psum5),0)
'end if

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
if isnumeric(rateSum) then
	if int(rateSum) > 100 then
		rateSum = 100
	end if
end if

set RS1 = nothing
set RS2 = nothing

dim strTRFontColor1
dim strTRFontColor2
dim strTRFontColor3
dim strTRFontColor4
dim strTRFontColor5
dim strTRFontColor6
dim strTRFontColor7
dim strTRFontColorSum

'70% 이하일 때 적색 표시를 위한 처리
strTRBgClr1 = "black"
strTRFontColor1 = "white"
if isnumeric(rate1) then
	if int(rate1) < 90 then
		strTRFontColor1 = "white"
		strTRBgClr1 = "darkred"
	end if
	if int(rate1) < 80 then
		strTRFontColor1 = "white"
		strTRBgClr1 = "darkred"
	end if
	rate1 = rate1 & "%"
end if

strTRBgClr2 = "black"
strTRFontColor2 = "white"
if isnumeric(rate2) then
	if int(rate2) < 90 then
		strTRFontColor2 = "white"
		strTRBgClr2 = "darkred"
	end if
	if int(rate2) < 80 then
		strTRFontColor2 = "white"
		strTRBgClr2 = "darkred"
	end if
	rate2 = rate2 & "%"
end if

strTRBgClr3 = "black"
strTRFontColor3 = "white"
if isnumeric(rate3) then
	if int(rate3) < 90 then
		strTRFontColor3 = "white"
		strTRBgClr3 = "darkred"
	end if
	if int(rate3) < 80 then
		strTRFontColor3 = "white"
		strTRBgClr3 = "darkred"
	end if
	rate3 = rate3 & "%"
end if

strTRBgClr4 = "black"
strTRFontColor4 = "white"
if isnumeric(rate4) then
	if int(rate4) < 90 then
		strTRFontColor4 = "white"
		strTRBgClr4 = "darkred"
	end if
	if int(rate4) < 80 then
		strTRFontColor4 = "white"
		strTRBgClr4 = "darkred"
	end if
	rate4 = rate4 & "%"
end if

strTRBgClr5 = "black"
strTRFontColor5 = "white"
if isnumeric(rate5) then
	if int(rate5) < 90 then
		strTRFontColor5 = "white"
		strTRBgClr5 = "darkred"
	end if
	if int(rate5) < 80 then
		strTRFontColor5 = "white"
		strTRBgClr5 = "darkred"
	end if
	rate5 = rate5 & "%"
end if

strTRBgClr6 = "black"
strTRFontColor6 = "white"
if isnumeric(rate6) then
	if int(rate6) < 90 then
		strTRFontColor6 = "white"
		strTRBgClr6 = "darkred"
	end if
	if int(rate6) < 80 then
		strTRFontColor6 = "white"
		strTRBgClr6 = "darkred"
	end if
	rate6 = rate6 & "%"
end if

strTRBgClr7 = "black"
strTRFontColor7 = "white"
if isnumeric(rate7) then
	if int(rate7) < 90 then
		strTRFontColor7 = "white"
		strTRBgClr7 = "darkred"
	end if
	if int(rate7) < 80 then
		strTRFontColor7 = "white"
		strTRBgClr7 = "darkred"
	end if
	rate7 = rate7 & "%"
end if

strTRBgClrSum = "black"
strTRFontColorSum = "white"
if isnumeric(rateSum) then
	if int(rateSum) < 90 then
		strTRFontColorSum = "white"
		strTRBgClrSum = "darkred"
	end if
	if int(rateSum) < 80 then
		strTRFontColorSum = "white"
		strTRBgClrSum = "darkred"
	end if
	rateSum = rateSum & "%"
end if

'목표수량을 0으로 초기화
idlesum1 = 0
idlesum2 = 0
idlesum3 = 0
idlesum4 = 0
idlesum5 = 0
idlesum6 = 0
idlesum7 = 0

if s_Process = "PCBA" then
	idlesum1 = getIdleSum("PCBA1")
	idlesum2 = getIdleSum("PCBA2")
	idlesum3 = getIdleSum("PCBA3")
	idlesum4 = getIdleSum("PCBA4")
	idlesum5 = getIdleSum("PCBA5")
	idlesum6 = getIdleSum("PCBA6")
	idlesum7 = getIdleSum("PCBA7")
elseif s_Process = "CBOX" then
	idlesum1 = getIdleSum("CBOX1")
	idlesum2 = getIdleSum("CBOX2")
	idlesum3 = getIdleSum("CBOX3")
	idlesum4 = getIdleSum("CBOX4")
	idlesum5 = getIdleSum("CBOX5")
	idlesum6 = getIdleSum("CBOX6")
	idlesum7 = getIdleSum("CBOX7")
end if
idlesumTotal = idlesum1+idlesum2+idlesum3+idlesum4+idlesum5+idlesum6+idlesum7

idlesum1 = makeHMMSS(idlesum1)
idlesum2 = makeHMMSS(idlesum2)
idlesum3 = makeHMMSS(idlesum3)
idlesum4 = makeHMMSS(idlesum4)
idlesum5 = makeHMMSS(idlesum5)
idlesum6 = makeHMMSS(idlesum6)
idlesum7 = makeHMMSS(idlesum7)
idlesumTotal = makeHMMSS(idlesumTotal)
%>
<%

function makeHMMSS(nSec)
	dim hh
	dim mm
	dim ss
	hh = int(nSec / 3600)
	mm = int((nSec mod 3600)/60)
	ss = int((nSec mod 3600) mod 60)
	
	if (nSec mod 3600)/60 > 0 then
		mm = mm + 1
	end if
	if hh < 10 then
		hh = "0" & hh
	end if
	if mm < 10 then
		mm = "0" & mm
	end if
	if ss < 10 then
		ss = "0" & ss
	end if
	makeHMMSS = hh &":"& mm '&":"&ss
end function

'무작업시간 계산 
function getIdleSum(strLine)
	dim SQL
	dim RS1
	dim CNT1, CNT2
	dim arrIdle(50,2)
	dim nIdleSum
	
	dim oldLine_State_LS_State
	dim Line_State_LS_State
	dim oldLSL_Date
	dim LSL_Date
	
	dim chkRest1
	dim chkRest2
	dim chkRest3
	dim chkRest4
	
	dim RestInsideIdle
	dim arrRest(3,1)
	'근무 시간 8시20분~10시20분, 10시30분~12시30분, 13시10분~15시10분, 15시20분~17시20분, 17시40~20시40분
	arrRest(0,0)=date()&" 오전 10:20:00"
	arrRest(0,1)=date()&" 오전 10:30:00"
	arrRest(1,0)=date()&" 오후 12:30:00"
	arrRest(1,1)=date()&" 오후 1:10:00"
	arrRest(2,0)=date()&" 오후 3:10:00"
	arrRest(2,1)=date()&" 오후 3:20:00"
	arrRest(3,0)=date()&" 오후 5:20:00"
	arrRest(3,1)=date()&" 오후 5:40:00"
	
	
	CNT1 = 0
	set RS1 = Server.CreateObject("Adodb.RecordSet")
	SQL = "select Line_State_LS_State, LSL_Date from tbLine_State_Log Where Line_State_LS_Line = '"&strLine&"' and LSL_Date between '"&date()&"' and '"&dateadd("D",1,date())&"' order by LSL_Code asc"
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		
		Line_State_LS_State	= RS1("Line_State_LS_State")
		LSL_Date			= RS1("LSL_Date")
		
		'정상이 아니라면(무작업시작이라면 배열의 0열에 시간 저장)
		if Line_State_LS_State <> "정상" then
			if arrIdle(CNT1,0) = "" then
				arrIdle(CNT1,0) = LSL_Date
			end if
		'정상이라면(무작업시작이라면 배열의 0열에 시간 저장)
		elseif arrIdle(CNT1,0) <> "" then
			arrIdle(CNT1,1) = LSL_Date
			CNT1 = CNT1 + 1 '다음행으로 이동 
		end if
		
		'response.write DateDiff("m",oldLSL_Date,LSL_Date) &"<br>"

		oldLine_State_LS_State	= RS1("Line_State_LS_State")
		oldLSL_Date				= RS1("LSL_Date")
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
	

	'아직 무작업 상태라면,
	if oldLine_State_LS_State <> "정상" then
		arrIdle(CNT1,1) = now
	end if
	
	nIdleSum = 0
	for CNT1 = 0 to ubound(arrIdle)
		'response.write arrIdle(CNT1,0) &"_"&arrIdle(CNT1,1)&"<Br>"
		if arrIdle(CNT1,0) <> "" and arrIdle(CNT1,1) <> "" then
			'response.write "CNT1:"&CNT1&"    arrIdle0:"&arrIdle(CNT1,0)&"    arrIdle1:"&arrIdle(CNT1,1)&"<br>"
			'response.write datediff("s",arrIdle(CNT1,0),arrIdle(CNT1,1))&"<bR>"
			
			RestInsideIdle = 0
			'쉬는 시간을 거친 수 만큼 쉬는 시간 차감
			for CNT2 = 0 to ubound(arrRest)
			'response.write arrIdle(CNT1,0) &"___"& arrRest(CNT2,0) &"___"&datediff("s",arrIdle(CNT1,0),arrRest(CNT2,0))&"<br>"
			'response.write arrIdle(CNT1,0) &"___"& arrRest(CNT2,1) &"___"&datediff("s",arrIdle(CNT1,0),arrRest(CNT2,1))&"<br>"
			'response.write arrIdle(CNT1,1) &"___"& arrRest(CNT2,0) &"___"&datediff("s",arrIdle(CNT1,1),arrRest(CNT2,0))&"<br>"
			'response.write arrIdle(CNT1,1) &"___"& arrRest(CNT2,1) &"___"&datediff("s",arrIdle(CNT1,1),arrRest(CNT2,1))&"<br><br>"
				chkRest1 = datediff("s",arrIdle(CNT1,0),arrRest(CNT2,0))
			
				chkRest2 = datediff("s",arrIdle(CNT1,0),arrRest(CNT2,1))
				chkRest3 = datediff("s",arrIdle(CNT1,1),arrRest(CNT2,0))
				chkRest4 = datediff("s",arrIdle(CNT1,1),arrRest(CNT2,1))
				
			    '무무무무무무무'
							'쉬쉬쉬쉬쉬쉬쉬'
				if chkRest1 >= 0 and chkRest2 >= 0 and chkRest3 >= 0 and chkRest4 >= 0 then
				'무무무무무무무'
					'쉬쉬쉬쉬쉬쉬쉬'
				elseif chkRest1 >= 0 and chkRest2 >= 0 and chkRest3 <= 0 and chkRest4 >= 0 then
					arrIdle(CNT1,1) = arrRest(CNT2,0)
					'무무무무무무무'
				'쉬쉬쉬쉬쉬쉬쉬'
				elseif chkRest1 <= 0 and chkRest2 >= 0 and chkRest3 <= 0 and chkRest4 <= 0 then
					arrIdle(CNT1,0) = arrRest(CNT2,1)
							'무무무무무무무'
				'쉬쉬쉬쉬쉬쉬쉬'
				elseif chkRest1 <= 0 and chkRest2 <= 0 and chkRest3 <= 0 and chkRest4 <= 0 then
				'무무무무무무무무무무무무무무'
					'쉬쉬쉬쉬쉬쉬쉬'
				elseif chkRest1 >= 0 and chkRest2 >= 0 and chkRest3 <= 0 and chkRest4 <= 0 then
					RestInsideIdle = datediff("s",arrRest(CNT2,0),arrRest(CNT2,1))
					'무무무무무무무'
				'쉬쉬쉬쉬쉬쉬쉬쉬쉬쉬쉬쉬쉬쉬'
				elseif chkRest1 <= 0 and chkRest2 >= 0 and chkRest3 <= 0 and chkRest4 >= 0 then
					RestInsideIdle = datediff("s",arrIdle(CNT1,0),arrIdle(CNT1,1))
				end if
			next
		
			nIdleSum = nIdleSum + datediff("s",arrIdle(CNT1,0),arrIdle(CNT1,1)) - RestInsideIdle
		end if
	next
	
	getIdleSum = nIdleSum
end function

'목표수량 계산 함수
function getTargetQty(strLine, calcNow)
	dim CNT1
	
	dim SQL
	dim RS1
	dim tQty
	
	dim B_D_No
	dim oldB_D_No
	dim lenDiff
 
	dim PSP_Count
	dim BP_PPH
	dim PSP_ST
	dim ChangeOverHead
	
	dim accSec
	dim accQty
	
	

	set RS1 = server.CreateObject("ADODB.RecordSet")

	tQty = 0
	accSec = 0
	accQty = 0
	SQL = ""
	SQL = SQL & "select "
	SQL = SQL & "	t1.BOM_Sub_BS_D_No, "
	SQL = SQL & "	t1.PSP_Count, "
	SQL = SQL & "	BP_PPH = isnull((select top 1 t2.BP_PPH from tbBOM_PPH t2 where t2.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No),0) "
	SQL = SQL & "from tbProcess_State_Plan t1 "
	SQL = SQL & "where t1.PSP_Line = '"&strLine&"' and t1.PSP_Work_Date = '"&s_Work_Date&"' "
	SQL = SQL & "order by PSP_Code asc "
	RS1.Open SQL,sys_DBCon
	
	'계획테이블 루프를 돈다. 인자(라인, 날짜)
	
	ChangeOverHead		= 0
	oldB_D_No	= ""
	do until RS1.Eof 
		
		'모델/옵션 체인지 체크
		B_D_No = RS1("BOM_Sub_BS_D_No")
		'제일 처음은 패스
		if oldB_D_No <> "" then
			'옵션번호 지우기
			if isnumeric(left(B_D_No,4)) then '6871계열이라면
				B_D_No = left(B_D_No,10)
			else
				B_D_No = left(B_D_No,9)
			end if
			
			'혹시 유사모델체인지인지 모델체인지 인지 확인하자
			for CNT1 = 0 to ubound(arrSimilar) - 1
				arrSimilarDetail = split(arrSimilar(CNT1),"$")
				
				'만약 유사모델 리스트에 있다면, 대표파트넘버로 바꾼다
				if instr(arrSimilarDetail(1), "-"&B_D_No&"-") > 0 then
					B_D_No = arrSimilarDetail(0)
				end if
			next
			
			'설사 같은 모델이라도 기본으로 옵션은 바뀌겠지.
			ChangeOverHead = 1
			if B_D_No <> oldB_D_No then
				ChangeOverHead = 4
			end if
		end if
		oldB_D_No = B_D_No
		
		PSP_Count 		= RS1("PSP_Count") '계획수량
		BP_PPH			= RS1("BP_PPH")
		if BP_PPH = 0 then
			BP_PPH = 300
		end if
		
		PSP_ST	= cint(3600 / BP_PPH) '개당 생산시간
		
		'response.write "LINE:"&strLine&"accSec:"&accSec&"    accQty:"&accQty&"     PSP_Count:"&PSP_Count&"     PSP_ST:"&PSP_ST&"     OH:"&ChangeOverHead&"<Br>"
		'이번 레코드의 총 생산필요시간을 accSec에 누적 / 총 계획수량을 accQty에 누적 / 오버헤드 반영
		accSec = accSec + (PSP_Count * PSP_ST) + (ChangeOverHead*60)
		accQty = accQty + PSP_Count
		
		'response.write "accSec:"&accSec&"    accQty:"&accQty&"     PSP_Count:"&PSP_Count&"     PSP_ST:"&PSP_ST&"     OH:"&ChangeOverHead&"<Br>"
		
		'누적된 필요시간이 현재시간과 같다면
		if calcNow = accSec then
			
			getTargetQty = accQty '지금까지의 수량을 반환
			exit do
		'현재시간을 지나간다면
		elseif calcNow < accSec then
			'정확한 수량을 계산하기 위해...
			accSec = accSec - (PSP_Count * PSP_ST) '마지막으로 누적한 생산필요시간을 뺀다.
			accQty = accQty - PSP_Count '마지막으로 누적한 계획수량을 뺀다.
			
			do until calcNow < accSec '최대생산가능수량까지
				accSec = accSec + PSP_ST '생산시간을 더한다
				accQty = accQty + 1 '수량을 하나씩 늘린다 
			loop
			
			getTargetQty = accQty '지금까지의 수량을 반환
			exit do
		end if
		RS1.MoveNext
	loop
	RS1.Close
	
	set RS1 = nothing
	
	if isnull(getTargetQty) or getTargetQty = "" then
		getTargetQty = accQty
	end if
end function

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
strHTML += "<table width=100% cellpadding=0 cellspacing=3 bgcolor='white' style='color:white;font-size:65px;text-align:center;font-weight:bold'>";
strHTML += "<col width=180px></col>";
strHTML += "<col width=210px></col>";
strHTML += "<col width=210px></col>";
strHTML += "<col width=210px></col>";
strHTML += "<col width=215px></col>";
strHTML += "<col width=210px></col>";
strHTML += "<col></col>";
strHTML += "<tr height=100px bgcolor=skyblue style='color:navy'>";
strHTML += "	<td>LINE</td>";
strHTML += "	<td>계획</td>";
strHTML += "	<td>목표</td>";
strHTML += "	<td>실적</td>";
strHTML += "	<td>달성율</td>";
strHTML += "	<td>무작업</td>";
strHTML += "	<td>상태</td>";
strHTML += "</tr>";
strHTML += "<tr height=100px bgcolor=<%=strTRBgClr1%> style='color:<%=strTRFontColor1%>'>";
strHTML += "	<td><%=strLineTitle%>M1</td>";
strHTML += "	<td align=right><%=psum1%></td>";
strHTML += "	<td align=right><%=tsum1%></td>";
strHTML += "	<td align=right>"+nSum1+"</td>";
strHTML += "	<td align=right><%=rate1%></td>";
strHTML += "	<td align=center ><%=idlesum1%></td>";
strHTML += "	<td bgcolor=<%=strBgClr1%>><%=strLineState1%></td>";
strHTML += "</tr>";
strHTML += "<tr height=100px bgcolor=<%=strTRBgClr2%> style='color:<%=strTRFontColor2%>'>";
strHTML += "	<td><%=strLineTitle%>M2</td>";
strHTML += "	<td align=right><%=psum2%></td>";
strHTML += "	<td align=right><%=tsum2%></td>";
strHTML += "	<td align=right>"+nSum2+"</td>";
strHTML += "	<td align=right><%=rate2%></td>";
strHTML += "	<td align=center ><%=idlesum2%></td>";
strHTML += "	<td bgcolor=<%=strBgClr2%>><%=strLineState2%></td>";
strHTML += "</tr>";
strHTML += "<tr height=100px bgcolor=<%=strTRBgClr3%> style='color:<%=strTRFontColor3%>'>";
strHTML += "	<td><%=strLineTitle%>M3</td>";
strHTML += "	<td align=right><%=psum3%></td>";
strHTML += "	<td align=right><%=tsum3%></td>";
strHTML += "	<td align=right>"+nSum3+"</td>";
strHTML += "	<td align=right><%=rate3%></td>";
strHTML += "	<td align=center ><%=idlesum3%></td>";
strHTML += "	<td bgcolor=<%=strBgClr3%>><%=strLineState3%></td>";
strHTML += "</tr>";
strHTML += "<tr height=100px bgcolor=<%=strTRBgClr4%> style='color:<%=strTRFontColor4%>'>";
strHTML += "	<td><%=strLineTitle%>M4</td>";
strHTML += "	<td align=right><%=psum4%></td>";
strHTML += "	<td align=right><%=tsum4%></td>";
strHTML += "	<td align=right>"+nSum4+"</td>";
strHTML += "	<td align=right><%=rate4%></td>";
strHTML += "	<td align=center ><%=idlesum4%></td>";
strHTML += "	<td bgcolor=<%=strBgClr4%>><%=strLineState4%></td>";
strHTML += "</tr>";
strHTML += "<tr height=100px bgcolor=<%=strTRBgClr5%> style='color:<%=strTRFontColor5%>'>";
strHTML += "	<td><%=strLineTitle%>M5</td>";
strHTML += "	<td align=right><%=psum5%></td>";
strHTML += "	<td align=right><%=tsum5%></td>";
strHTML += "	<td align=right>"+nSum5+"</td>";
strHTML += "	<td align=right><%=rate5%></td>";
strHTML += "	<td align=center ><%=idlesum5%></td>";
strHTML += "	<td bgcolor=<%=strBgClr5%>><%=strLineState5%></td>";
strHTML += "</tr>";
strHTML += "<tr height=100px bgcolor=<%=strTRBgClrSum%> style='color:<%=strTRFontColorSum%>'>";
strHTML += "	<td>SUM</td>";
strHTML += "	<td align=right><%=psum1+psum2+psum3+psum4+psum5%></td>";
strHTML += "	<td align=right><%=tsum1+tsum2+tsum3+tsum4+tsum5%></td>";
strHTML += "	<td align=right>"+nSumTotal+"</td>";
strHTML += "	<td align=right><%=rateSum%></td>";
strHTML += "	<td align=center><%=idlesumTotal%></td>";
strHTML += "	<td bgcolor=<%=strBgClrSum%>>&nbsp;</td>";
strHTML += "</tr>";
strHTML += "<tr height=300px bgcolor=black>";
strHTML += "	<td colspan=7>&nbsp;</td>";
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
		location.href='mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=PCBA&s_Work_Date=<%=s_Work_Date%>'
<%
	else
%>
		location.href='mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=CBOX&s_Work_Date=<%=s_Work_Date%>'
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
		location.href='mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=PCBA&s_Work_Date=<%=s_Work_Date%>'
<%
	else
%>
		location.href='mtr_Content_Process_State_all_Record.asp?s_Multi_YN=Y&s_Process=CBOX&s_Work_Date=<%=s_Work_Date%>'
<%
	end if
%>
	}
	else
	{
		setTimeout("fRun()",30);
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
		setTimeout("fRun()",30);
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


	
	