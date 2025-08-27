<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim RS1
dim RS2
dim SQL

dim s_Work_Date
dim s_Line

s_Work_Date = date()
s_Line = request("s_Line")

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

calcNow = left(FormatDateTime(now(),4),2)*60 + right(FormatDateTime(now(),4),2)
calcNow = calcNow * 60
calcNow = getRestedCalcNow(calcNow)
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

'
function MakePlanTable()
	dim CNT1
		
	dim SQL
	dim RS1
	dim tQty
	
	dim BS_D_No
	dim B_D_No
	dim oldB_D_No
	dim lenDiff
	
	dim PSP_Count
	dim BP_PPH
	dim PSP_ST
	dim ChangeOverHead
	
	dim accSec
	dim accQty
	
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
	
	'최초 생산시간 구하기
	SQL = "select minPRD_Input_Time = min(PRD_Input_Time) from tbPWS_Raw_Data "
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and PRD_Line = '"&s_Line&"' and "&vbcrlf
	SQL = SQL & "	PRD_Input_Date is not null"&vbcrlf
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		calcPRD_Start = 30000
	else
		calcPRD_Start = (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		if calcPRD_Start < 30000 then '8시 20분 이전에 생산이 시작되었다면
			calcPRD_Start = 30000 '그냥 8시 20분으로 보정
		end if
	end if
	RS1.Close
	
	'계획 루프돌기
	tQty = 0
	accSec = 0
	accQty = 0
	SQL = ""
	SQL = SQL & "select "
	SQL = SQL & "	t1.BOM_Sub_BS_D_No, "
	SQL = SQL & "	t1.PSP_Count, "
	SQL = SQL & "	BP_PPH = isnull((select top 1 t2.BP_PPH from tbBOM_PPH t2 where t2.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No),0) "
	SQL = SQL & "from tbProcess_State_Plan t1 "
	SQL = SQL & "where t1.PSP_Line = '"&s_Line&"' and t1.PSP_Work_Date = '"&s_Work_Date&"' "
	SQL = SQL & "order by PSP_Code asc "
	RS1.Open SQL,sys_DBCon
	
	ChangeOverHead	= 0
	oldB_D_No		= ""
	do until RS1.Eof 
		
		'모델/옵션 체인지 체크
		B_D_No	= RS1("BOM_Sub_BS_D_No")
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
		
		BS_D_No		= RS1("BOM_Sub_BS_D_No")
		PSP_Count 	= RS1("PSP_Count") '계획수량
		BP_PPH		= RS1("BP_PPH")
		if BP_PPH = 0 then
			BP_PPH = 300
		end if
		
		PSP_ST	= cint(3600 / BP_PPH) '개당 생산시간
		
		'이번 레코드의 총 생산필요시간을 accSec에 누적 / 총 계획수량을 accQty에 누적 / 오버헤드 반영
		accSec = accSec + (PSP_Count * PSP_ST) + (ChangeOverHead*60)
		accQty = accQty + PSP_Count
		
		'누적된 필요시간이 2시간과 같다면
		if accSec = 2*60*60 then
			accSec = 0
			accQty = 0
			strPlanTable = strPlanTable & BS_D_No & "$" & PSP_Count & "//"
			strPlanTable = strPlanTable & "휴식" & "$" & "0" & "//"
			splitYN = "N"
			RS1.MoveNext
		'2시간을 지나간다면
		elseif 2*60*60 < accSec then
			'정확한 수량을 계산하기 위해...
			accSec = accSec - (PSP_Count * PSP_ST) '마지막으로 누적한 생산필요시간을 뺀다.
			accQty = accQty - PSP_Count '마지막으로 누적한 계획수량을 뺀다.
			
			do until 2*60*60 < accSec '최대생산가능수량까지
				accSec = accSec + PSP_ST '생산시간을 더한다
				accQty = accQty + 1 '수량을 하나씩 늘린다 
				accQtyPre = accQtyPre + 1
			loop
			
			accSec = 0
			accQty = 0
			strPlanTable = strPlanTable & BS_D_No & "$" & accQtyPre & "//"
			strPlanTable = strPlanTable & "휴식" & "$" & "0" & "//"
			PSP_Count = PSP_Count - accQtyPre
			splitYN = "Y"
		else
			strPlanTable = strPlanTable & BS_D_No & "$" & PSP_Count & "//"
			splitYN = "N"
			RS1.MoveNext
		end if
		
		
	loop
	RS1.Close

	set RS1 = nothing
end function

if isnull(getTargetQty) or getTargetQty = "" then
	getTargetQty = accQty
end if
RS1.Close




set RS1 = nothing

response.end
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
strHTML += "<tr bgcolor=green style='color:white'>";
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
strHTML += "	<td><%=arrRemain(CNT1)%></td>";	//계획-누적수량
strHTML += "	<td><%=left(arrPWS_Opt_Data(CNT1,3),2)%>:<%=right(arrPWS_Opt_Data(CNT1,3),2)%> - <%=left(arrPWS_Opt_Data(CNT1,4),2)%>:<%=right(arrPWS_Opt_Data(CNT1,4),2)%></td>";	//작업시간
strHTML += "	<td align=right><%=arrRndRate(CNT1)%>&nbsp;</td>";	
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

function reload_handle()
{
	location.reload();
}

</script>


<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


