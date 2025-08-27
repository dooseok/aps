<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim flag_YN
dim nST
dim nSTCalc
dim nRAWCalc
dim rndRate
dim rndRate2
dim nRemain

dim CurrentRecord

dim s_Work_Date
dim s_Line
dim s_Jaje_YN

's_Work_Date = request("s_Work_Date")
s_Work_Date = date()
s_Line = request("s_Line")
s_Jaje_YN = request("s_Jaje_YN")

dim RS1
dim SQL

'반복문에 사용하기 위한 변수 선언
dim CNT1
dim CNT2
dim CNT3

dim oldPWS_Opt_Data_5

dim arrPWS_Merged_Data			'병합 용 배열
dim arrPWS_Opt_Data				'출력하기 위해 최적화 된 배열

dim arrPWS_Raw_Data				'실적 정보 배열
dim arrPWS_Raw_Data_Shrinked	'실적 정보 배열 (계획 수량 차감 후)

dim arrPWS_Rest_Data(3,8)		'휴식 정보 배열

dim arrPWS_Plan_Data			'계획 정보 배열
dim arrPWS_Plan_Sub				'계획 정보 배열(Sub)

'0 파트넘버
'1 수량(계획)
'2 누적수량
'3 시작시각
'4 종료시각
'5 비고
'6 실적수량
'7 실적시작
'8 실적종료

dim strPWS_Plan_Data_0
dim strPWS_Plan_Data_1
dim strPWS_Plan_Data_2
dim strPWS_Plan_Data_3
dim strPWS_Plan_Data_4
dim strPWS_Plan_Data_5
dim strPWS_Plan_Data_6
dim strPWS_Plan_Data_7
dim strPWS_Plan_Data_8
dim arrPWS_Plan_Data_0
dim arrPWS_Plan_Data_1
dim arrPWS_Plan_Data_2
dim arrPWS_Plan_Data_3
dim arrPWS_Plan_Data_4
dim arrPWS_Plan_Data_5
dim arrPWS_Plan_Data_6
dim arrPWS_Plan_Data_7
dim arrPWS_Plan_Data_8

dim strPWS_Plan_Sub_0
dim strPWS_Plan_Sub_1
dim strPWS_Plan_Sub_2
dim arrPWS_Plan_Sub_0
dim arrPWS_Plan_Sub_1
dim arrPWS_Plan_Sub_2

set RS1 = Server.CreateObject("ADODB.RecordSet")

'실적 정보의 배열화
arrPWS_Raw_Data = getPWS_Raw_Data(s_Work_Date, s_Line)


'휴식 정보의 배열화
arrPWS_Rest_Data(0,0) = "무작업"	'1st 휴식
arrPWS_Rest_Data(0,1) = "0"
arrPWS_Rest_Data(0,2) = "0"
arrPWS_Rest_Data(0,3) = "1020"
arrPWS_Rest_Data(0,4) = "1030"
arrPWS_Rest_Data(0,5) = "휴식"
arrPWS_Rest_Data(0,6) = "0"
arrPWS_Rest_Data(0,7) = "1020"
arrPWS_Rest_Data(0,8) = "1030"
arrPWS_Rest_Data(1,0) = "무작업"	'2nd 휴식
arrPWS_Rest_Data(1,1) = "0"
arrPWS_Rest_Data(1,2) = "0"
arrPWS_Rest_Data(1,3) = "1230"
arrPWS_Rest_Data(1,4) = "1320"
arrPWS_Rest_Data(1,5) = "휴식"
arrPWS_Rest_Data(1,6) = "0"
arrPWS_Rest_Data(1,7) = "1230"
arrPWS_Rest_Data(1,8) = "1310"
arrPWS_Rest_Data(2,0) = "무작업"	'3rd 휴식
arrPWS_Rest_Data(2,1) = "0"
arrPWS_Rest_Data(2,2) = "0"
arrPWS_Rest_Data(2,3) = "1510"
arrPWS_Rest_Data(2,4) = "1520"
arrPWS_Rest_Data(2,5) = "휴식"
arrPWS_Rest_Data(2,6) = "0"
arrPWS_Rest_Data(2,7) = "1510"
arrPWS_Rest_Data(2,8) = "1520"
arrPWS_Rest_Data(3,0) = "무작업"	'4th 휴식
arrPWS_Rest_Data(3,1) = "0"
arrPWS_Rest_Data(3,2) = "0"
arrPWS_Rest_Data(3,3) = "1720"
arrPWS_Rest_Data(3,4) = "1740"
arrPWS_Rest_Data(3,5) = "휴식"
arrPWS_Rest_Data(3,6) = "0"
arrPWS_Rest_Data(3,7) = "1720"
arrPWS_Rest_Data(3,8) = "1740"

'계획 정보의 배열화
SQL = ""
SQL = SQL & "select BOM_Sub_BS_D_No, PSP_Start, PSP_End, PSP_Desc, PSP_Count, PSP_Sub_YN, PSP_Sub_Start, PSP_Sub_End  from tbProcess_State_Plan "
SQL = SQL & "where PSP_Work_Date = '"&s_Work_Date&"' and PSP_Line = '"&s_Line&"' order by PSP_Start "
'SQL = SQL & "group by BOM_Sub_BS_D_No"

RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strPWS_Plan_Data_0	= strPWS_Plan_Data_0	& RS1("BOM_Sub_BS_D_No")			& "|/|"
	strPWS_Plan_Data_1	= strPWS_Plan_Data_1	& "0"								& "|/|"
	strPWS_Plan_Data_2	= strPWS_Plan_Data_2	& "0"								& "|/|"	
	strPWS_Plan_Data_3	= strPWS_Plan_Data_3	& int(RS1("PSP_Start")) + 2000		& "|/|"
	strPWS_Plan_Data_4	= strPWS_Plan_Data_4	& int(RS1("PSP_End")) + 2000		& "|/|"
	strPWS_Plan_Data_5	= strPWS_Plan_Data_5	& RS1("PSP_Desc")					& "|/|"
	strPWS_Plan_Data_6	= strPWS_Plan_Data_6	& RS1("PSP_Count")					& "|/|"
	strPWS_Plan_Data_7	= strPWS_Plan_Data_7	& RS1("PSP_Start")					& "|/|"
	strPWS_Plan_Data_8	= strPWS_Plan_Data_8	& RS1("PSP_End")					& "|/|"
	
	strPWS_Plan_Sub_0	= strPWS_Plan_Sub_0		& RS1("PSP_Sub_YN")					& "|/|"
	strPWS_Plan_Sub_1	= strPWS_Plan_Sub_1		& RS1("PSP_Sub_Start")				& "|/|"
	strPWS_Plan_Sub_2	= strPWS_Plan_Sub_2		& RS1("PSP_Sub_End")				& "|/|"
	RS1.MoveNext
loop
RS1.Close

if strPWS_Plan_Data_0 <> "" then
	arrPWS_Plan_Data_0	= split(strPWS_Plan_Data_0,"|/|")
	arrPWS_Plan_Data_1	= split(strPWS_Plan_Data_1,"|/|")
	arrPWS_Plan_Data_2	= split(strPWS_Plan_Data_2,"|/|")
	arrPWS_Plan_Data_3	= split(strPWS_Plan_Data_3,"|/|")
	arrPWS_Plan_Data_4	= split(strPWS_Plan_Data_4,"|/|")
	arrPWS_Plan_Data_5	= split(strPWS_Plan_Data_5,"|/|")
	arrPWS_Plan_Data_6	= split(strPWS_Plan_Data_6,"|/|")
	arrPWS_Plan_Data_7	= split(strPWS_Plan_Data_7,"|/|")
	arrPWS_Plan_Data_8	= split(strPWS_Plan_Data_8,"|/|")
	
	arrPWS_Plan_Sub_0	= split(strPWS_Plan_Sub_0,"|/|")
	arrPWS_Plan_Sub_1	= split(strPWS_Plan_Sub_1,"|/|")
	arrPWS_Plan_Sub_2	= split(strPWS_Plan_Sub_2,"|/|")
	
	redim arrPWS_Plan_Data(ubound(arrPWS_Plan_Data_0)-1, 8)
	redim arrPWS_Plan_Sub(ubound(arrPWS_Plan_Data_0)-1, 2)
	
	for CNT1 = 0 to ubound(arrPWS_Plan_Data_0)-1
		arrPWS_Plan_Data(CNT1,0)	= arrPWS_Plan_Data_0(CNT1)
		arrPWS_Plan_Data(CNT1,1)	= arrPWS_Plan_Data_1(CNT1)
		arrPWS_Plan_Data(CNT1,2)	= arrPWS_Plan_Data_2(CNT1)
		arrPWS_Plan_Data(CNT1,3)	= arrPWS_Plan_Data_3(CNT1)
		arrPWS_Plan_Data(CNT1,4)	= arrPWS_Plan_Data_4(CNT1)
		arrPWS_Plan_Data(CNT1,5)	= arrPWS_Plan_Data_5(CNT1)
		arrPWS_Plan_Data(CNT1,6)	= arrPWS_Plan_Data_6(CNT1)
		arrPWS_Plan_Data(CNT1,7)	= arrPWS_Plan_Data_7(CNT1)
		arrPWS_Plan_Data(CNT1,8)	= arrPWS_Plan_Data_8(CNT1)
		
		arrPWS_Plan_Sub(CNT1,0)		= arrPWS_Plan_Sub_0(CNT1)
		arrPWS_Plan_Sub(CNT1,1)		= arrPWS_Plan_Sub_1(CNT1)
		arrPWS_Plan_Sub(CNT1,2)		= arrPWS_Plan_Sub_2(CNT1)
		
		'SUB PCB 계획이고 투입 완료 되었다면
		if arrPWS_Plan_Sub(CNT1,0) = "Y" and len(arrPWS_Plan_Sub(CNT1,1)) = 4 and len(arrPWS_Plan_Sub(CNT1,2)) = 4 then
			
			arrPWS_Plan_Data(CNT1,1) = arrPWS_Plan_Data(CNT1,6)
			arrPWS_Plan_Data(CNT1,2) = arrPWS_Plan_Data(CNT1,6)
			arrPWS_Plan_Data(CNT1,3) = arrPWS_Plan_Sub(CNT1,1)
			arrPWS_Plan_Data(CNT1,4) = arrPWS_Plan_Sub(CNT1,2)
			arrPWS_Plan_Data(CNT1,5) = "raw"
			
		end if
	next
else
	redim arrPWS_Plan_Data(0, 8)
end if

'for CNT1 = 0 to ubound(arrPWS_Plan_Data)
'	response.write arrPWS_Plan_Data(CNT1,0) & "/" & arrPWS_Plan_Data(CNT1,6) & "<Br>"
'next
'response.write "<Br>"

'동일 모델이 연이어 있을 때, Merging처리
CNT3 = 0
for CNT1=1 to ubound(arrPWS_Plan_Data)
	'직전의 모델과 동일하다면
	
	'두 계획의 시간간격이 10분 미만이라면.
	'response.write int(left(arrPWS_Plan_Data(CNT1,7),2)*60+right(arrPWS_Plan_Data(CNT1,7),2)) - int(left(arrPWS_Plan_Data(CNT1-1,8),2)*60+right(arrPWS_Plan_Data(CNT1-1,8),2))
	if arrPWS_Plan_Data(CNT1,0) <> "" and arrPWS_Plan_Data(CNT1-1,0) = arrPWS_Plan_Data(CNT1,0) and int(left(arrPWS_Plan_Data(CNT1,7),2)*60+right(arrPWS_Plan_Data(CNT1,7),2)) - int(left(arrPWS_Plan_Data(CNT1-1,8),2)*60+right(arrPWS_Plan_Data(CNT1-1,8),2)) < 10 then
		'직전 레코드에 수량만 합하면 된다.
		arrPWS_Plan_Data(CNT1-1,6)	= int(arrPWS_Plan_Data(CNT1-1,6)) + int(arrPWS_Plan_Data(CNT1,6))
		arrPWS_Plan_Data(CNT1-1,8)	= arrPWS_Plan_Data(CNT1,8)
		
		'지금 있는 칸으로 한칸씩 당겨오기
		for CNT2 = CNT1 to ubound(arrPWS_Plan_Data) - 1
			arrPWS_Plan_Data(CNT2,0) = arrPWS_Plan_Data(CNT2+1,0)
			arrPWS_Plan_Data(CNT2,1) = arrPWS_Plan_Data(CNT2+1,1)
			arrPWS_Plan_Data(CNT2,2) = arrPWS_Plan_Data(CNT2+1,2)
			arrPWS_Plan_Data(CNT2,3) = arrPWS_Plan_Data(CNT2+1,3)
			arrPWS_Plan_Data(CNT2,4) = arrPWS_Plan_Data(CNT2+1,4)
			arrPWS_Plan_Data(CNT2,5) = arrPWS_Plan_Data(CNT2+1,5)
			arrPWS_Plan_Data(CNT2,6) = arrPWS_Plan_Data(CNT2+1,6)
			arrPWS_Plan_Data(CNT2,7) = arrPWS_Plan_Data(CNT2+1,7)
			arrPWS_Plan_Data(CNT2,8) = arrPWS_Plan_Data(CNT2+1,8)
		next
		arrPWS_Plan_Data(CNT2,0) = ""
		CNT1=CNT1-1 '레코드가 하나씩 당겨진 셈이므로, CNT1을 다시 수행하기 위해 차감한다.
		CNT3=CNT3+1
	end if
next

'for CNT1 = 0 to ubound(arrPWS_Plan_Data)
'	response.write arrPWS_Plan_Data(CNT1,0) & "/" & arrPWS_Plan_Data(CNT1,6) & "<Br>"
'next
'response.write "<Br>"

dim strRawFlag1
strRawFlag1 = "-"
dim strRawFlag2
strRawFlag2 = "-"

dim nRemainPlan

dim flag_YN2

for CNT1=0 to ubound(arrPWS_Plan_Data)

	if arrPWS_Plan_Data(CNT1,6) < 1 then
		flag_YN = "Y"
	end if
	
	'실적 배열을 루핑
	flag_YN = "N"
	
	if flag_YN = "N" then
		for CNT2=0 to ubound(arrPWS_Raw_Data)
			'계획과 실적의 파트넘버가 동일하다면 그리고 한번이상 적용 받지 않아야 하고 
			if arrPWS_Raw_Data(CNT2,0) = arrPWS_Plan_Data(CNT1,0) and flag_YN = "N" and (instr(strRawFlag1,"-"&arrPWS_Raw_Data(CNT2,0)&"-") < 1 or instr(strRawFlag2,"-"&arrPWS_Raw_Data(CNT2,3)&"-") < 1) then
				
				
				if isnumeric(arrPWS_Raw_Data(CNT2,6)) and isnumeric(arrPWS_Plan_Data(CNT1,6)) then
					arrPWS_Raw_Data(CNT2,6) = int(arrPWS_Raw_Data(CNT2,6)) + int(arrPWS_Plan_Data(CNT1,6))
				elseif isnumeric(arrPWS_Plan_Data(CNT1,6)) then
					arrPWS_Raw_Data(CNT2,6) = int(arrPWS_Plan_Data(CNT1,6))
				elseif isnumeric(arrPWS_Raw_Data(CNT2,6)) then
					arrPWS_Raw_Data(CNT2,6) = int(arrPWS_Raw_Data(CNT2,6))
				end if
				
				arrPWS_Raw_Data(CNT2,7) = arrPWS_Plan_Data(CNT1,7)
				arrPWS_Raw_Data(CNT2,8) = arrPWS_Plan_Data(CNT1,8)
				
				'실적이 계획보다 많으면 다음에 적용을 받을 수 있도록 한다.
				'if int(arrPWS_Plan_Data(CNT1,6)) >= int(arrPWS_Raw_Data(CNT2,1)) then
					strRawFlag1 = strRawFlag1 & arrPWS_Raw_Data(CNT2,0) & "-"
					strRawFlag2 = strRawFlag2 & arrPWS_Raw_Data(CNT2,3) & "-"
				'end if
				
				flag_YN = "Y"
			end if
		next
	end if
	if flag_YN = "Y" then
		arrPWS_Plan_Data(CNT1,0) = ""
	end if 
next

'for CNT1 = 0 to ubound(arrPWS_Plan_Data)
'	response.write arrPWS_Plan_Data(CNT1,0) & "/" & arrPWS_Plan_Data(CNT1,6) & "<Br>"
'next
'response.write "<Br>"

'for CNT1 = 0 to ubound(arrPWS_Raw_Data)
'	response.write arrPWS_Raw_Data(CNT1,0) & "/" & arrPWS_Raw_Data(CNT1,1) & "<Br>"
'next
'response.write "<Br>"




'redim arrPWS_Raw_Data_Shrinked(ubound(arrPWS_Raw_Data)-CNT3+1,8)

'CNT2 = 0
'for CNT1=0 to ubound(arrPWS_Raw_Data)
'	if arrPWS_Raw_Data(CNT1,0) <> "" then
'		arrPWS_Raw_Data_Shrinked(CNT2,0) = arrPWS_Raw_Data(CNT1,0)
'		arrPWS_Raw_Data_Shrinked(CNT2,1) = arrPWS_Raw_Data(CNT1,1)
'		arrPWS_Raw_Data_Shrinked(CNT2,2) = arrPWS_Raw_Data(CNT1,2)
'		arrPWS_Raw_Data_Shrinked(CNT2,3) = arrPWS_Raw_Data(CNT1,3)
'		arrPWS_Raw_Data_Shrinked(CNT2,4) = arrPWS_Raw_Data(CNT1,4)
'		arrPWS_Raw_Data_Shrinked(CNT2,5) = arrPWS_Raw_Data(CNT1,5)
'		arrPWS_Raw_Data_Shrinked(CNT2,6) = arrPWS_Raw_Data(CNT1,6)
'		arrPWS_Raw_Data_Shrinked(CNT2,7) = arrPWS_Raw_Data(CNT1,7)
'		arrPWS_Raw_Data_Shrinked(CNT2,8) = arrPWS_Raw_Data(CNT1,8)
'		CNT2 = CNT2 + 1
'	end if
'next

'시작시간 기준으로 정렬
arrPWS_Merged_Data = Merging_Array(arrPWS_Raw_Data, arrPWS_Plan_Data, arrPWS_Rest_Data)
arrPWS_Opt_Data = QuickSort(arrPWS_Merged_Data, lbound(arrPWS_Merged_Data), ubound(arrPWS_Merged_Data), 3, "ASC")

'for CNT1 = 0 to ubound(arrPWS_Merged_Data)
'response.write arrPWS_Merged_Data(CNT1,0) & " / "
'response.write arrPWS_Merged_Data(CNT1,1) & " / "
'response.write arrPWS_Merged_Data(CNT1,2) & " / "
'response.write arrPWS_Merged_Data(CNT1,3) & " / "
'response.write arrPWS_Merged_Data(CNT1,4) & " / "
'response.write arrPWS_Merged_Data(CNT1,5) & " / "
'response.write arrPWS_Merged_Data(CNT1,6) & " / "
'response.write arrPWS_Merged_Data(CNT1,7) & " / "
'response.write arrPWS_Merged_Data(CNT1,8) & "<Br>"
'next

CurrentRecord = 0
for CNT1 = 0 to ubound(arrPWS_Opt_Data)
	if arrPWS_Opt_Data(CNT1,5) = "raw" then
		CurrentRecord = CNT1	
	end if
next


%>

<script language="javascript">
var strHTML = "";
strHTML += "<table width=100% cellpadding=0 cellspacing=1 bgcolor='white' style='color:white;font-size:37px;text-align:center;font-weight:bold'>";
strHTML += "<col></col>";
strHTML += "<col width=150px></col>";
strHTML += "<col width=150px></col>";
strHTML += "<col width=150px></col>";
strHTML += "<col width=300px></col>";
strHTML += "<col width=150px></col>";
<%
dim nRawCount
nRawCount = 0

dim sumPlan
dim sumRaw
sumPlan = 0
sumRaw = 0
for CNT1=0 to ubound(arrPWS_Opt_Data)

	if arrPWS_Opt_Data(CNT1,0) = "" then
	

	elseif arrPWS_Opt_Data(CNT1,0) = "무작업" and cint(replace(arrPWS_Opt_Data(CNT1,3),":","")) > cint(replace(FormatDateTime(now(),4),":","")) then
		
	elseif arrPWS_Opt_Data(CNT1,0) = "무작업" then
%>
strHTML += "<tr bgcolor=black>";
strHTML += "	<td colspan=4><%if arrPWS_Opt_Data(CNT1,5)="휴식" then%>휴&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;식<%else%><%=arrPWS_Opt_Data(CNT1,5)%><%end if%></td>";
strHTML += "	<td><%=left(arrPWS_Opt_Data(CNT1,3),2)%>:<%=right(arrPWS_Opt_Data(CNT1,3),2)%> - <%=left(arrPWS_Opt_Data(CNT1,4),2)%>:<%=right(arrPWS_Opt_Data(CNT1,4),2)%></td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "</tr>";
<%
	else
		'실적에서 가져온 배열이라면
		if arrPWS_Opt_Data(CNT1,5) = "raw" then
			nRawCount = nRawCount + 1
		
			if CurrentRecord = int(CNT1) then
%>
strHTML += "<tr bgcolor=red style='color:white'>";
<%
			else
%>
strHTML += "<tr bgcolor=black>";
<%
			end if
%>
strHTML += "	<td><span style='cursor:hand' onclick=\"javascript:Pop_Print('<%=arrPWS_Opt_Data(CNT1,0)%>');\"><%=arrPWS_Opt_Data(CNT1,0)%></span></td>";	//작업순서
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,6)%></td>";	//계획수량
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,1)%></td>"; //실적수량
<%
			nRemain = 0
			flag_YN = "N"
			for CNT3 = 0 to CNT1
				if arrPWS_Opt_Data(CNT3,0) = arrPWS_Opt_Data(CNT1,0) then
					if isnumeric(arrPWS_Opt_Data(CNT3,1)) then
						sumRaw = int(sumRaw) + int(arrPWS_Opt_Data(CNT3,1))
					end if
					
					if isnumeric(arrPWS_Opt_Data(CNT3,6)) then
						sumPlan = int(sumPlan) + int(arrPWS_Opt_Data(CNT3,6))
					end if
					
					flag_YN = "Y"

				end if
			next
			
			
			
			if flag_YN = "Y" then
				nRemain = sumPlan - sumRaw
				sumPlan = 0
				sumRaw	= 0
			else
				if isnumeric(arrPWS_Opt_Data(CNT1,6)) and isnumeric(arrPWS_Opt_Data(CNT1,2)) then
					nRemain = int(arrPWS_Opt_Data(CNT1,6)) - int(arrPWS_Opt_Data(CNT1,2))
					if nRemain < 0 then
						'nRemain = 0
					end if
				else
					'nRemain = 0
				end if
			end if
			
%>
strHTML += "	<td><%=nRemain%></td>";	//계획-누적수량
strHTML += "	<td><%=left(arrPWS_Opt_Data(CNT1,3),2)%>:<%=right(arrPWS_Opt_Data(CNT1,3),2)%> - <%=left(arrPWS_Opt_Data(CNT1,4),2)%>:<%=right(arrPWS_Opt_Data(CNT1,4),2)%></td>";	//작업시간
<%			
			'실적수량 x st
			SQL = "select top 1 BS_ST from tbBOM_Sub where BS_D_No = '"&arrPWS_Opt_Data(CNT1,0)&"'"
			
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then
				nST = "8"
			else
				if RS1(0) = 0 or isnull(RS1(0)) then
					nST = "8"
				else
					nST = RS1(0)
				end if
			end if
			RS1.Close
						
			nSTCalc = arrPWS_Opt_Data(CNT1,1) * nST
			
			nRAWCalc = ((int(left(arrPWS_Opt_Data(CNT1,4),2)) * 60 + int(right(arrPWS_Opt_Data(CNT1,4),2))) - (int(left(arrPWS_Opt_Data(CNT1,3),2)) * 60 + int(right(arrPWS_Opt_Data(CNT1,3),2)))) * 60
		
			if nRAWCalc = 0 then
				nRAWCalc = 1
			end if
			
			if isnumeric(nRAWCalc) and nRAWCalc > 0 then
				rndRate = round(nSTCalc/nRAWCalc*100,0)
				if rndRate > 100 then
					rndRate = 100
				end if
				rndRate = rndRate & "%"
			else
				rndRate = "-"
			end if	
			
			if int(arrPWS_Opt_Data(CNT1,1)) < 10 then
				rndRate = "-"
			end if
			
%>
strHTML += "	<td align=right><%=rndRate%>&nbsp;</td>";	
strHTML += "</tr>";
<%
		'아직 실적은 없는 계획만 있는 레코드 라면
		else
%>
strHTML += "<tr bgcolor=black>";
strHTML += "	<td><span style='cursor:hand' onclick=\"javascript:Pop_Print('<%=arrPWS_Opt_Data(CNT1,0)%>');\"><%=arrPWS_Opt_Data(CNT1,0)%></span></td>";	//작업순서
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,6)%></td>";	//계획수량
strHTML += "	<td>0</td>";	//실적수량
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,6)%></td>";	//잔량
strHTML += "	<td><%=left(arrPWS_Opt_Data(CNT1,7),2)%>:<%=right(arrPWS_Opt_Data(CNT1,7),2)%> - <%=left(arrPWS_Opt_Data(CNT1,8),2)%>:<%=right(arrPWS_Opt_Data(CNT1,8),2)%></td>"; //계획시간
strHTML += "	<td align=center>-</td>";	//달성율
strHTML += "</tr>";
<%
		end if
	end if
	
	oldPWS_Opt_Data_5 = arrPWS_Opt_Data(CNT1,5)
next

set RS1 = Nothing
%>

parent.idContent.innerHTML = strHTML;
var nScroll = <%=nRawCount%>*43-600;
if(nScroll > 0)
	parent.scrollTo(0,nScroll);
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


</script>

<%
function Merging_Array(arr1, arr2, arr3)
	dim CNT1
	dim CNT2
	
	dim arrOpt	
	redim arrOpt(ubound(arr1)+ubound(arr2)+ubound(arr3)+2,8)
	dim arrTemp(8)

	for CNT1=0 to ubound(arr1)
		arrOpt(CNT1,0) = arr1(CNT1,0)
		arrOpt(CNT1,1) = arr1(CNT1,1)
		arrOpt(CNT1,2) = arr1(CNT1,2)
		arrOpt(CNT1,3) = arr1(CNT1,3)
		arrOpt(CNT1,4) = arr1(CNT1,4)
		arrOpt(CNT1,5) = arr1(CNT1,5)
		arrOpt(CNT1,6) = arr1(CNT1,6)
		arrOpt(CNT1,7) = arr1(CNT1,7)
		arrOpt(CNT1,8) = arr1(CNT1,8)
	next

	for CNT1=0 to ubound(arr2)
		arrOpt(CNT1+ubound(arr1)+1,0) = arr2(CNT1,0)
		arrOpt(CNT1+ubound(arr1)+1,1) = arr2(CNT1,1)
		arrOpt(CNT1+ubound(arr1)+1,2) = arr2(CNT1,2)
		arrOpt(CNT1+ubound(arr1)+1,3) = arr2(CNT1,3)
		arrOpt(CNT1+ubound(arr1)+1,4) = arr2(CNT1,4)
		arrOpt(CNT1+ubound(arr1)+1,5) = arr2(CNT1,5)
		arrOpt(CNT1+ubound(arr1)+1,6) = arr2(CNT1,6)
		arrOpt(CNT1+ubound(arr1)+1,7) = arr2(CNT1,7)
		arrOpt(CNT1+ubound(arr1)+1,8) = arr2(CNT1,8)
	next

	for CNT1=0 to ubound(arr3)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,0) = arr3(CNT1,0)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,1) = arr3(CNT1,1)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,2) = arr3(CNT1,2)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,3) = arr3(CNT1,3)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,4) = arr3(CNT1,4)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,5) = arr3(CNT1,5)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,6) = arr3(CNT1,6)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,7) = arr3(CNT1,7)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2)+2,8) = arr3(CNT1,8)
	next
	
	Merging_Array = arrOpt

end function
%>

<%
function Ordering_Array(arr1, arr2, arr3)
	dim CNT1
	dim CNT2
	
	dim arrOpt	
	redim arrOpt(ubound(arr1)+ubound(arr2)+ubound(arr3),8)
	dim arrTemp(8)

	for CNT1=0 to ubound(arr1)	
	
		arrOpt(CNT1,0) = arr1(CNT1,0)
		arrOpt(CNT1,1) = arr1(CNT1,1)
		arrOpt(CNT1,2) = arr1(CNT1,2)
		arrOpt(CNT1,3) = arr1(CNT1,3)
		arrOpt(CNT1,4) = arr1(CNT1,4)
		arrOpt(CNT1,5) = arr1(CNT1,5)
		arrOpt(CNT1,6) = arr1(CNT1,6)
		arrOpt(CNT1,7) = arr1(CNT1,7)
		arrOpt(CNT1,8) = arr1(CNT1,8)
	next

	for CNT1=0 to ubound(arr2)
		arrOpt(CNT1+ubound(arr1),0) = arr2(CNT1,0)
		arrOpt(CNT1+ubound(arr1),1) = arr2(CNT1,1)
		arrOpt(CNT1+ubound(arr1),2) = arr2(CNT1,2)
		arrOpt(CNT1+ubound(arr1),3) = arr2(CNT1,3)
		arrOpt(CNT1+ubound(arr1),4) = arr2(CNT1,4)
		arrOpt(CNT1+ubound(arr1),5) = arr2(CNT1,5)
		arrOpt(CNT1+ubound(arr1),6) = arr2(CNT1,6)
		arrOpt(CNT1+ubound(arr1),7) = arr2(CNT1,7)
		arrOpt(CNT1+ubound(arr1),8) = arr2(CNT1,8)
	next

	for CNT1=0 to ubound(arr3)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),0) = arr3(CNT1,0)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),1) = arr3(CNT1,1)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),2) = arr3(CNT1,2)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),3) = arr3(CNT1,3)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),4) = arr3(CNT1,4)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),5) = arr3(CNT1,5)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),6) = arr3(CNT1,6)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),7) = arr3(CNT1,7)
		arrOpt(CNT1+ubound(arr1)+ubound(arr2),8) = arr3(CNT1,8)
	next

	for CNT1=0 to ubound(arrOpt,1)-1    '처음 배열부터 마지막 전까지
		for CNT2=CNT1+1 to ubound(arrOpt,1)   '두번째 배열부터 마지막까지
		'값을 비교합니다. > < 구분으로 오름차수 내림차순을 결정합니다..
			if clng(arrOpt(CNT1,3)) > clng(arrOpt(CNT2,3)) then
				arrTemp(0) = arrOpt(CNT1,0)
				arrTemp(1) = arrOpt(CNT1,1)
				arrTemp(2) = arrOpt(CNT1,2)
				arrTemp(3) = arrOpt(CNT1,3)
				arrTemp(4) = arrOpt(CNT1,4)
				arrTemp(5) = arrOpt(CNT1,5)
				arrTemp(6) = arrOpt(CNT1,6)
				arrTemp(7) = arrOpt(CNT1,7)
				arrTemp(8) = arrOpt(CNT1,8)
				arrOpt(CNT1,0) = arrOpt(CNT2,0)
				arrOpt(CNT1,1) = arrOpt(CNT2,1)
				arrOpt(CNT1,2) = arrOpt(CNT2,2)
				arrOpt(CNT1,3) = arrOpt(CNT2,3)
				arrOpt(CNT1,4) = arrOpt(CNT2,4)
				arrOpt(CNT1,5) = arrOpt(CNT2,5)
				arrOpt(CNT1,6) = arrOpt(CNT2,6)
				arrOpt(CNT1,7) = arrOpt(CNT2,7)
				arrOpt(CNT1,8) = arrOpt(CNT2,8)
				arrOpt(CNT2,0) = arrTemp(0)
				arrOpt(CNT2,1) = arrTemp(1)
				arrOpt(CNT2,2) = arrTemp(2)
				arrOpt(CNT2,3) = arrTemp(3)
				arrOpt(CNT2,4) = arrTemp(4)
				arrOpt(CNT2,5) = arrTemp(5)
				arrOpt(CNT2,6) = arrTemp(6)
				arrOpt(CNT2,7) = arrTemp(7)
				arrOpt(CNT2,8) = arrTemp(8)
			end if
		next
	next

	Ordering_Array = arrOpt

end function
%>

<%
function QuickSort(vec,loBound,hiBound,SortField,SortDir)
  '==--------------------------------------------------------==
  '== Sort a multi dimensional array on SortField            ==
  '==                                                        ==
  '== This procedure is adapted from the algorithm given in: ==
  '==    ~ Data Abstractions & Structures using C++ by ~     ==
  '==    ~ Mark Headington and David Riley, pg. 586    ~     ==
  '== Quicksort is the fastest array sorting routine for     ==
  '== unordered arrays.  Its big O is n log n                ==
  '==                                                        ==
  '== Parameters:                                            ==
  '== vec       - array to be sorted                         ==
  '== SortField - The field to sort on (1st dimension value) ==
  '== loBound and hiBound are simply the upper and lower     ==
  '==   bounds of the array's "row" dimension. It's probably ==
  '==   easiest to use the LBound and UBound functions to    ==
  '==   set these.                                           ==
  '== SortDir   - ASC, ascending; DESC, Descending           ==
  '==--------------------------------------------------------==
  if not (hiBound - loBound = 0) then
      Dim pivot(),loSwap,hiSwap,temp,counter
      Redim pivot (Ubound(vec,2))
      SortDir = UCase(SortDir)

      '== Two items to sort
      if hiBound - loBound = 1 then
        if (SortDir = "ASC") then
            if FormatCompare(vec(loBound,SortField),vec(hiBound,SortField)) > FormatCompare(vec(hiBound,SortField),vec(loBound,SortField)) then Call SwapRows(vec,hiBound,loBound)
        else
            if FormatCompare(vec(loBound,SortField),vec(hiBound,SortField)) < FormatCompare(vec(hiBound,SortField),vec(loBound,SortField)) then Call SwapRows(vec,hiBound,loBound)
        end if
      End If

      '== Three or more items to sort
      For counter = 0 to Ubound(vec,2)
        pivot(counter) = vec(int((loBound + hiBound) / 2),counter)
        vec(int((loBound + hiBound) / 2),counter) = vec(loBound,counter)
        vec(loBound,counter) = pivot(counter)
      Next

      loSwap = loBound + 1
      hiSwap = hiBound

      do
        '== Find the right loSwap
        if (SortDir = "ASC") then
            while loSwap < hiSwap and FormatCompare(vec(loSwap,SortField),pivot(SortField)) <= FormatCompare(pivot(SortField),vec(loSwap,SortField))
              loSwap = loSwap + 1
            wend
        else
            while loSwap < hiSwap and FormatCompare(vec(loSwap,SortField),pivot(SortField)) >= FormatCompare(pivot(SortField),vec(loSwap,SortField))
              loSwap = loSwap + 1
            wend
        end if
        '== Find the right hiSwap
        if (SortDir = "ASC") then
            while FormatCompare(vec(hiSwap,SortField),pivot(SortField)) > FormatCompare(pivot(SortField),vec(hiSwap,SortField))
              hiSwap = hiSwap - 1
            wend
        else
            while FormatCompare(vec(hiSwap,SortField),pivot(SortField)) < FormatCompare(pivot(SortField),vec(hiSwap,SortField))
              hiSwap = hiSwap - 1
            wend
        end if
        '== Swap values if loSwap is less then hiSwap
        if loSwap < hiSwap then Call SwapRows(vec,loSwap,hiSwap)
      loop while loSwap < hiSwap

      For counter = 0 to Ubound(vec,2)
        vec(loBound,counter) = vec(hiSwap,counter)
        vec(hiSwap,counter) = pivot(counter)
      Next

      '== Recursively call function .. the beauty of Quicksort
        '== 2 or more items in first section
        if loBound < (hiSwap - 1) then Call QuickSort(vec,loBound,hiSwap-1,SortField,SortDir)
        '== 2 or more items in second section
        if hiSwap + 1 < hibound then Call QuickSort(vec,hiSwap+1,hiBound,SortField,SortDir)
  end if
  
  QuickSort = vec
end function  'QuickSort
%>

<%
Sub SwapRows(ary,row1,row2)
  '==------------------------------------------==
  '== This proc swaps two rows of an array     ==
  '==------------------------------------------==

  Dim x,tempvar
  For x = 0 to Ubound(ary,2)
    tempvar = ary(row1,x)
    ary(row1,x) = ary(row2,x)
    ary(row2,x) = tempvar
  Next
End Sub  'SwapRows
%>

<%
function FormatCompare(sOne,sTwo)
  '==------------------------------------------==
  '==  Checks sOne & sTwo, returns sOne as a   ==
  '==  Numeric if both pass isNumeric, if not  ==
  '==  returns sOne as a string.               ==
  '==------------------------------------------==

    if (isNumeric(Trim(sOne)) AND isNumeric(Trim(sTwo))) then
        FormatCompare = CDbl(Trim(sOne))
    else
        FormatCompare = Trim(sOne)
    end if
end function
%>

<%
Sub PrintArray(vec,loRow,hiRow,markCol)
  '==------------------------------------------==
  '== Print out an array  Highlight the column ==
  '==  whose number matches param markCol      ==
  '==------------------------------------------==

  Dim ColNmbr,RowNmbr
  Response.Write "<table border=""1"" cellspacing=""0"">"
  For RowNmbr = loRow to hiRow
    Response.Write "<tr>"
    For ColNmbr = 0 to (Ubound(vec,2) - 1)
      If ColNmbr = markCol then
        Response.Write "<td bgcolor=""FFFFCC"">"
      Else
        Response.Write "<td>"
      End If
      Response.Write vec(RowNmbr,ColNmbr) & "</td>"
    Next
    Response.Write "</tr>"
  Next
  Response.Write "</table>"
End Sub  'PrintArray
%>
<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	
	