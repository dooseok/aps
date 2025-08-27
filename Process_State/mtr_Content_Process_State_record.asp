<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim RS1
dim RS2
dim SQL
dim CNT1
dim CNT2
dim s_Work_Date
dim s_Line

s_Work_Date = date()
's_Work_Date = "2015-08-24"
s_Line		= Request("s_Line")

dim calcPRD_Start

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

set RS1 = server.CreateObject("ADODB.RecordSet")

'-------------------------------------
'생산시간 테이블 만들기
'-------------------------------------
'최초 생산시간 구하기
SQL = "select minPRD_Input_Time = min(PRD_Input_Time) from tbPWS_Raw_Data "
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and PRD_Line = '"&s_Line&"' and "&vbcrlf
SQL = SQL & "	PRD_Input_Date is not null "&vbcrlf
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
	calcPRD_Start = 30000
elseif isnull(RS1(0)) then
	calcPRD_Start = 30000
else
	calcPRD_Start = (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
	if calcPRD_Start < 30000 then '8시 20분 이전에 생산이 시작되었다면
		calcPRD_Start = 30000 '그냥 8시 20분으로 보정
	end if
end if
RS1.Close
dim strWorkTimeTable
dim arrWorkTimeTable

'쉬는 시간 중이라면, 쉬는 시간 시작 상태로 고정
if calcPRD_Start >= 500*60 and calcPRD_Start <= 620*60 then '8시 20분 ~ 10시 20분
	strWorkTimeTable = cstr(620 - int(calcPRD_Start/60)) & "/120/120/120/9999"
elseif calcPRD_Start > 620*60 and calcPRD_Start <= 630*60 then '10시 20분~10시 30분 
	strWorkTimeTable = "0/120/120/120/9999"
elseif calcPRD_Start > 630*60 and calcPRD_Start <= 750*60 then '10시 30분~12시 30분 
	strWorkTimeTable = "0/"&cstr(750 - int(calcPRD_Start/60))&"/120/120/9999"
elseif calcPRD_Start > 750*60 and calcPRD_Start <= 790*60 then '12시 30분~13시 10분 
	strWorkTimeTable = "0/0/120/120/9999"
elseif calcPRD_Start > 790*60 and calcPRD_Start <= 910*60 then '13시 10분~15시 10분 
	strWorkTimeTable = "0/0/"&cstr(910 - int(calcPRD_Start/60))&"/120/9999"
elseif calcPRD_Start > 910*60 and calcPRD_Start <= 920*60 then '15시 10분~15시 20분 
	strWorkTimeTable = "0/0/0/120/9999"
elseif calcPRD_Start > 920*60 and calcPRD_Start <= 1040*60 then '15시 20분~17시 20분 
	strWorkTimeTable = "0/0/0/"&cstr(1040 - int(calcPRD_Start/60))&"/9999"
elseif calcPRD_Start > 1040*60 and calcPRD_Start <= 1060*60 then '17시 20분~40분
	strWorkTimeTable = "0/0/0/0/9999"
elseif calcPRD_Start > 1060*60 then '17시 40분~
	strWorkTimeTable = "0/0/0/0/9999"
end if

arrWorkTimeTable = split(strWorkTimeTable,"/")

'-------------------------------------
'-------------------------------------




'-------------------------------------
'생산실적 배열화 
'-------------------------------------
dim strPRD_PartNo
dim strCntPRD_Code
dim arrPRD_PartNo
dim arrCntPRD_Code
SQL = "select PRD_PartNo, cntPRD_Code = count(PRD_Code) from tbPWS_Raw_Data where PRD_Line = '"&s_Line&"' and (PRD_ICT_Date = '"&s_Work_Date&"' or PRD_Input_Date = '"&s_Work_Date&"') group by PRD_PartNo"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strPRD_PartNo	= strPRD_PartNo		&RS1("PRD_PartNo")	&"/"
	strCntPRD_Code	= strCntPRD_Code	&RS1("cntPRD_Code")	&"/"
	RS1.MoveNext
loop
RS1.Close
if len(strPRD_PartNo) > 0 then
	strPRD_PartNo	= left(strPRD_PartNo,	len(strPRD_PartNo)-1)
	strCntPRD_Code	= left(strCntPRD_Code,	len(strCntPRD_Code)-1)
end if 
arrPRD_PartNo	= split(strPRD_PartNo,	"/")
arrCntPRD_Code	= split(strCntPRD_Code,	"/")
'-------------------------------------
'-------------------------------------





'-------------------------------------
'계획 DB 가져오기 
'-------------------------------------
dim strBS_D_No
dim strPSP_Count
dim strPSP_ST
dim arrBS_D_No
dim arrPSP_Count
dim arrPSP_ST
dim BP_PPH
SQL = ""
SQL = SQL & "select "
SQL = SQL & "	t1.BOM_Sub_BS_D_No, "
SQL = SQL & "	t1.PSP_Count, "
SQL = SQL & "	BP_PPH = isnull((select top 1 t2.BP_PPH from tbBOM_PPH t2 where t2.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No),0) "
SQL = SQL & "from tbProcess_State_Plan t1 "
SQL = SQL & "where t1.PSP_Line = '"&s_Line&"' and t1.PSP_Work_Date = '"&s_Work_Date&"' "
SQL = SQL & "order by PSP_Code asc "
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strBS_D_No		= strBS_D_No	&RS1("BOM_Sub_BS_D_No")	&"/"
	strPSP_Count	= strPSP_Count	&RS1("PSP_Count")		&"/"
	
	if BP_PPH = 0 then
		BP_PPH = 300
	end if
	strPSP_ST		= strPSP_ST		&cint(3600 / BP_PPH)	&"/"
	RS1.MoveNext
loop
RS1.Close
if len(strBS_D_No) > 0 then
	strBS_D_No		= left(strBS_D_No,	len(strBS_D_No)-1)
	strPSP_Count	= left(strPSP_Count,len(strPSP_Count)-1)
	strPSP_ST		= left(strPSP_ST,	len(strPSP_ST)-1)
end if 
arrBS_D_No		= split(strBS_D_No,		"/")
arrPSP_Count	= split(strPSP_Count,	"/")
arrPSP_ST		= split(strPSP_ST,		"/")
'-------------------------------------
'-------------------------------------

dim accSec
dim arrMaster(200,6)
dim CNT3
dim CNT4
dim CNT5
CNT3 = 0
CNT4 = 0

accSec = 0
dim accQty
accQty = 0
dim nOverHead
dim BS_D_No
dim oldBS_D_No
dim B_D_No
dim oldB_D_No

'-------------------------------------
'계획 테이블 만들기(계획DB와 생산시간Table 조합하기)
'-------------------------------------
'계획 DB를 루핑한다.

for CNT1 = 0 to ubound(arrBS_D_No) '이번 모델
	
	'모델체인지 오버헤드 가져오기
	if CNT1 > 0 then
		BS_D_No = arrBS_D_No(CNT1)
		oldBS_D_No = arrBS_D_No(CNT1-1)
		'옵션번호 지우기
		if isnumeric(left(BS_D_No,4)) then '6871계열이라면
			B_D_No = left(BS_D_No,10)
		else
			B_D_No = left(BS_D_No,9)
		end if
		if isnumeric(left(oldBS_D_No,4)) then '6871계열이라면
			oldB_D_No = left(oldBS_D_No,10)
		else
			oldB_D_No = left(oldBS_D_No,9)
		end if
		
		'혹시 유사모델체인지인지 모델체인지 인지 확인하자
		for CNT5 = 0 to ubound(arrSimilar) - 1
			arrSimilarDetail = split(arrSimilar(CNT5),"$")
			
			'만약 유사모델 리스트에 있다면, 대표파트넘버로 바꾼다
			if instr(arrSimilarDetail(1), "-"&B_D_No&"-") > 0 then
				B_D_No = arrSimilarDetail(0)
			end if
			'만약 유사모델 리스트에 있다면, 대표파트넘버로 바꾼다
			if instr(arrSimilarDetail(1), "-"&oldB_D_No&"-") > 0 then
				oldB_D_No = arrSimilarDetail(0)
			end if
		next
		
		'설사 같은 모델이라도 기본으로 옵션은 바뀌겠지.
		if B_D_No = oldB_D_No then
			nOverHead = 1
		else
			nOverHead = 4
		end if
		accSec = accSec + nOverHead * 60 '모델체인지 반영
	end if
	
	for CNT2 = 0 to arrPSP_Count(CNT1) '총 수량만큼 한개씩 올리면서 루핑
		
		accQty = accQty + 1 '생산수량 증가
		accSec = accSec + int(arrPSP_ST(CNT1)) '작업소요시간을 누적하여 증가
		'response.write CNT3 &"<br>"
		if int(accSec) <= int(arrWorkTimeTable(CNT3)*60) then '작업소요시간이 이번작업구간을 초과하지 않았다면,
			'계속 진행
		else '초과했다면 루프 빠져나감
			CNT3 = CNT3 + 1							'다음 작업구간으로 이동 
			arrMaster(CNT4,0) = arrBS_D_No(CNT1)	'테이블<파트넘버
			arrMaster(CNT4,1) = accQty				'테이블<초과당시의 계획수량
			CNT4 = CNT4 + 1
			arrMaster(CNT4,0) = "휴식"				'테이블<휴식시간
			arrMaster(CNT4,1) = 0					'테이블<휴식시간
			CNT4 = CNT4 + 1
			accQty	= 0							'누적생산 초기화
			accSec	= 0							'누적소요시간 초기화
			
		end if		
	next
	arrMaster(CNT4,0) = arrBS_D_No(CNT1)	'파트넘버
	arrMaster(CNT4,1) = accQty - 1			'초과당시의 계획수량
	CNT4 = CNT4 + 1
	accQty = 0	'누적생산 초기화
next
'-------------------------------------
'-------------------------------------

bottomPos = CNT4 - 1




'-------------------------------------
'마스터테이블 만들기(계획+실적)
'-------------------------------------
'마스터테이블 루핑 

dim bottomPos
for CNT1 = 0 to ubound(arrPRD_PartNo) '실적테이블 루핑 
	for CNT2 = 0 to ubound(arrMaster) '계획테이블 루핑 
		
		if arrPRD_PartNo(CNT1) = arrMaster(CNT2,0) then '같은 파트넘버를 찾은 경우
			
			if int(arrCntPRD_Code(CNT1)) <= int(arrMaster(CNT2,1)) then 		'계획수량 이하 생산중
				arrMaster(CNT2,2) = arrCntPRD_Code(CNT1) 	'계획 옆에 기록
				exit for									'루핑종료
			else	'마스터수량을 초과
				arrMaster(CNT2,2) = arrMaster(CNT2,1)		'계획에 일단 계획과 동일하게 기록
				arrCntPRD_Code(CNT1) = arrCntPRD_Code(CNT1) - arrMaster(CNT2,2) '적은 만큼 실적량을 지우고 다음루프
			end if
		end if
		
		
	next
next
'-------------------------------------
'-------------------------------------




'-------------------------------------
'가장 최근 생산 파트넘버를 찾아서 마스터 끝에서 부터 찾는다.
'-------------------------------------
dim lastPos
dim lastPartNo
SQL = ""
SQL = SQL & "select top 1 PRD_PartNo from tbPWS_Raw_Data where "
SQL = SQL & "	 PRD_Line = '"&s_Line&"' and "
SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
SQL = SQL & "order by "
SQL = SQL & " 	PRD_Input_Date desc, "
SQL = SQL & " 	PRD_Input_Time_Detail desc "
RS1.Open SQL,sys_DBCon

if not(RS1.Eof or RS1.Bof) then
	lastPartNo = RS1("PRD_PartNo")
	for CNT1 = 200 to 0	step -1
		if arrMaster(CNT1, 0) = lastPartNo then
			lastPos = CNT1
			exit for
		end if
	next
end if
RS1.Close
'-------------------------------------
'-------------------------------------



'-------------------------------------
'master정리
'0:파트넘버
'1:계획수량
'2:실적
'3:달성률
'------------------------------------- 
for CNT1=0 to ubound(arrMaster)
	if arrMaster(CNT1,0) <> "" then '마스터에 데이터가 있는 경우만 루핑
		if arrMaster(CNT1,2) = "" then
			arrMaster(CNT1,2) = 0
		end if
		
		if arrMaster(CNT1,1) = 0 then
			arrMaster(CNT1,3) = 0
		else
			arrMaster(CNT1,3) = cstr(round((arrMaster(CNT1,2)/arrMaster(CNT1,1)*100),0))&"%"
		end if
	end if
	'response.write arrMaster(CNT1,0) &"<BR>"
next
'-------------------------------------
'-------------------------------------


response.write now()
%>

<script language="javascript">
var strHTML = "";

strHTML += "<table width=100% cellpadding=0 cellspacing=1 bgcolor='white' style='color:white;font-size:37px;text-align:center;font-weight:bold'>";
strHTML += "<col width=350px></col>";
strHTML += "<col width=200px></col>";
strHTML += "<col width=200px></col>";
strHTML += "<col width=200px></col>";
strHTML += "<col></col>";
<%


for CNT1=0 to 200
	if arrMaster(CNT1,0) = "휴식" then
%>
strHTML += "<tr bgcolor=black>";
strHTML += "	<td colspan=5>휴&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;식</td>";
strHTML += "</tr>";
<%			
	elseif arrMaster(CNT1,0) = "" then
%>
strHTML += "<tr bgcolor=black>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "</tr>";
<%		
	else
		
		if lastPos <> "" and CNT1 = lastPos then
%>		
strHTML += "<tr bgcolor=green>";
<%	
		else
%>
strHTML += "<tr bgcolor=black>";
<%
		end if
%>
strHTML += "	<td><%=arrMaster(CNT1,0)%></td>";
strHTML += "	<td><%=arrMaster(CNT1,1)%></td>";	//계획수량
strHTML += "	<td><%=arrMaster(CNT1,2)%></td>";	//실적수량
strHTML += "	<td><%=arrMaster(CNT1,1)-arrMaster(CNT1,2)%></td>";	//잔량
strHTML += "	<td align=right><%=arrMaster(CNT1,3)%><img src='/img/blank' width=130px height=1px></td>";	//달성률
strHTML += "</tr>";
<%
	end if
next
%>
strHTML += "</table>";


parent.idContent.innerHTML = strHTML;
<%
if lastPos = "" then
	lastPos = bottomPos
	
end if
%>
var nScroll = <%=lastPos%>*43-650;
if(nScroll > 0)
	parent.scrollTo(0,nScroll);


function reload_handle()
{
	location.reload();
}

</script>


<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


