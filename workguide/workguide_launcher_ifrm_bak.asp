<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
'변수 선언!
dim RS1
dim SQL
dim arrRecordSet

dim WI_ProcessName
dim strWI_ImageFileName(15)
dim arrWI_ImageFileName

dim CNT1, CNT2
dim strWG_Pos
dim strWG_ResX
dim strWG_ResY
dim strWG_MCDelay
dim strWG_SlideDelay
dim strWG_Auto_YN
dim arrWG_Pos
dim arrWG_ResX
dim arrWG_ResY
dim arrWG_MCDelay
dim arrWG_SlideDelay
dim arrWG_Auto_YN

dim s_PRD_Line

dim strPrePartNo
dim arrPrePartNo
dim strIDX
dim arrIDX
dim strSlideCNT
dim arrSlideCNT

'이미지 리사이징
dim myImg
dim iWidth(15)
dim iheight(15)

dim strImageURL(15)

dim PRD_PartNo
dim arrPRD_PartNo(15)

dim nDelayUnitValue

dim dbtime
dim mctime

dim AMPM
dim arrHHMMSS

nDelayUnitValue = 10

'라인
s_PRD_Line		= request("s_PRD_Line")

'이전 파트넘버
strPrePartNo	= request("strPrePartNo")
arrPrePartNo	= split(strPrePartNo,",")

'슬라이드 번호 
strIDX			= request("strIDX")
arrIDX			= split(strIDX,",")

'몇번째 슬라이드인지
strSlideCNT		= request("strSlideCNT")
arrSlideCNT		= split(strSlideCNT,",")

set RS1 = server.CreateObject("ADODB.RecordSet")

'SQL = "insert into tbTest_setinterval (ts_Work,ts_Desc,ts_Now,ts_Diff) values ('WorkGuide','"&s_PRD_Line&"',getdate(),0)"
'sys_DBCon.execute(SQL)

'DB에서 모니터 설정값을 가져옴 
SQL = "select WG_Pos, WG_ResX, WG_ResY, WG_MCDelay, WG_SlideDelay, WG_Auto_YN from tbWorkguide where PRD_Line='"&s_PRD_Line&"' order by WG_Pos asc"
RS1.Open SQL,sys_DBCon

strWG_Pos			= ""
strWG_ResX			= ""
strWG_ResY			= ""
strWG_MCDelay		= ""
strWG_SlideDelay	= ""
strWG_Auto_YN		= ""
do until RS1.Eof
	strWG_Pos			= strWG_Pos			& RS1("WG_Pos")	& ","
	strWG_ResX			= strWG_ResX		& RS1("WG_ResX")& ","
	strWG_ResY			= strWG_ResY		& RS1("WG_ResY")& ","
	strWG_MCDelay		= strWG_MCDelay		& RS1("WG_MCDelay")& ","
	strWG_SlideDelay	= strWG_SlideDelay	& RS1("WG_SlideDelay")& ","
	strWG_Auto_YN		= strWG_Auto_YN		& RS1("WG_Auto_YN")& ","
	RS1.MoveNext
loop
RS1.Close
arrWG_Pos			= split(strWG_Pos,",")
arrWG_ResX			= split(strWG_ResX,",")
arrWG_ResY			= split(strWG_ResY,",")
'모델체인지 딜레이 
arrWG_MCDelay		= split(strWG_MCDelay,",")
'슬라이드 전환속도
arrWG_SlideDelay	= split(strWG_SlideDelay,",")
arrWG_Auto_YN		= split(strWG_Auto_YN,",")

'for CNT1 = 0 to 14 
'	SQL = "select top 1 PRD_PartNo from tbPWS_Raw_Data "&vbcrlf
'	SQL = SQL & "where "&vbcrlf
'	SQL = SQL & "	PRD_Line = '"&s_PRD_Line&"' and "&vbcrlf
'	SQL = SQL & "	DATEADD(s,"&-1*arrWG_MCDelay(CNT1)&",getdate()) "&vbcrlf
'	SQL = SQL & "	> "&vbcrlf
'	SQL = SQL & "	convert(datetime,convert(char(10),PRD_Input_date) + ' ' + convert(char(8),PRD_Input_Time_Detail)) "&vbcrlf
'	SQL = SQL & "order by PRD_Code desc "&vbcrlf
'	RS1.Open SQL,sys_DBCon
'	if RS1.Eof or RS1.Bof then
'		arrPRD_PartNo(CNT1) = ""
'	else
'		arrPRD_PartNo(CNT1) = RS1("PRD_PartNo")
'	end if
'	RS1.Close
'next

SQL = "select top 900 PRD_PartNo, PRD_Input_Date, PRD_Input_Time_Detail from tbPWS_Raw_Data "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	PRD_Line = '"&s_PRD_Line&"' and "&vbcrlf
SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "&vbcrlf
SQL = SQL & "order by PRD_Input_Date desc, PRD_Input_Time_Detail desc "&vbcrlf
RS1.Open SQL,sys_DBCon
if not(RS1.Eof or RS1.Bof) then
	arrRecordSet = RS1.GetRows()
end if
RS1.Close	
'response.write SQL
'각각의 모니터상의 모델체인지 딜레이에 맞춰서 각 모니터별 파트넘버를 가져옴 
CNT2 = 1
for CNT1 = 0 to ubound(arrRecordSet,2)
	dbtime = arrRecordSet(1,CNT1) &" "& arrRecordSet(2,CNT1)
	
	arrHHMMSS = split(formatDateTime(dateadd("s",-1*arrWG_MCDelay(CNT2),Time), vbLongTime),":")
	AMPM = left(arrHHMMSS(0),2)
	arrHHMMSS(0) = replace(arrHHMMSS(0),AMPM&" ","")
	
	'오전이고 12시라면
	if AMPM = "오전" and int(arrHHMMSS(0)) = 12 then
		arrHHMMSS(0) = 0
	elseif AMPM = "오후" and int(arrHHMMSS(0)) < 12 then
		arrHHMMSS(0) = int(arrHHMMSS(0)) + 12
	end if
	if arrHHMMSS(0) < 10 then
		 arrHHMMSS(0) = "0" & cstr(arrHHMMSS(0))
	end if

	mctime = date()&" "& cstr(arrHHMMSS(0))&":"&cstr(arrHHMMSS(1))&":"&cstr(arrHHMMSS(2))
	if datediff("s",dbtime,mctime) >= 0 then
	
		arrPRD_PartNo(CNT2) = arrRecordSet(0,CNT1)
		
		CNT1 = CNT1 - 1
		CNT2 = CNT2 + 1
		
	end if
	
	if CNT2 = 15 then
		exit for
	end if
next


SQL = ""
SQL = SQL & " select top 1 PRD_PartNo "
SQL = SQL & " from tbPWS_Raw_Data "
SQL = SQL & " where "
SQL = SQL & " 	PRD_Line='"&s_PRD_Line&"' and "
SQL = SQL & " 	PRD_Input_Date = '"&date()&"' and "
SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
SQL = SQL & " order by PRD_Input_Time_Detail desc "
RS1.Open SQL,sys_DBCon

PRD_PartNo = ""
if not(RS1.Eof or RS1.Bof) then
	PRD_PartNo = RS1("PRD_PartNo")
end if
RS1.Close

if PRD_PartNo = "" then
	SQL = ""
	SQL = SQL & " select top 1 PRD_PartNo "
	SQL = SQL & " from tbPWS_Raw_Data "
	SQL = SQL & " where "
	SQL = SQL & " 	PRD_Line='"&s_PRD_Line&"' and "
	SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
	SQL = SQL & " order by "
	SQL = SQL & " 	PRD_Input_Date desc, "
	SQL = SQL & " 	PRD_Input_Time_Detail desc "
	RS1.Open SQL,sys_DBCon
	if not(RS1.Eof or RS1.Bof) then
		PRD_PartNo = RS1("PRD_PartNo")
	end if
	RS1.Close
end if

'만약 모델체인지 속도가 0라면 그냥 가장 최근데이터를 가져옴
for CNT1 = 0 to 14 
	if arrPRD_PartNo(CNT1) = "" or int(arrWG_MCDelay(CNT1)) = 0 then
		arrPRD_PartNo(CNT1) = PRD_PartNo
	end if
next

'공정별 루프
for CNT1 = 0 to 14
	
	'해당파트넘버 및 해당공정에 일치하는 이미지파일목록을 가져온다.
	SQL = "select * from tbWorkGuideImage where WI_PartNo = '"&arrPRD_PartNo(CNT1)&"' and " &vbcrlf
	SQL = SQL & "WI_ProcessNumber = "&cint(CNT1)+1&" and WI_Temp_YN='N' and " &vbcrlf
	SQL = SQL & "(right(WI_ImageFileName,5) = '.jpeg' or right(WI_ImageFileName,4) in ('.jpg','.png','.gif')) " &vbcrlf
	SQL = SQL & " order by WI_ImageFileName asc" &vbcrlf
	RS1.Open SQL,sys_DBCon
	'response.write SQL
	strImageURL(CNT1) = ""
	
	'DB에 해당파트넘버에 대한 정보가 없다면
	if RS1.Eof or RS1.Bof then
		strImageURL(CNT1) = ""
	else
		WI_ProcessName = RS1("WI_ProcessName")
		do until RS1.Eof
			strImageURL(CNT1) = strImageURL(CNT1) & RS1("WI_ImageFileName") & "|%|"
			RS1.MoveNext
		loop
	end if
	RS1.Close
	
	strWI_ImageFileName(CNT1) = strImageURL(CNT1)
	arrWI_ImageFileName = split(strImageURL(CNT1),"|%|")
	'단일 파일인 경우 그대로 할당 
	if strImageURL(CNT1) <> "" then
		if ubound(arrWI_ImageFileName) = 1 then
			strImageURL(CNT1) = DefaultPath_workguide_img & arrPRD_PartNo(CNT1) & "\" & WI_ProcessName & "\" & arrWI_ImageFileName(0)
		'파일이 복수인 경우 
		else
			'이번에 모델체인지가 되었다면, 슬라이드 번호는 다시 처음으로
			if arrPRD_PartNo(CNT1) <> arrPrePartNo(CNT1) then  
				arrIDX(CNT1) = 0 '첫슬라이드로 셋팅
				arrSlideCNT(CNT1) = 0 '슬라이드 카운트도 처음으로
				
			'모델체인지가 된 것이 아니라면
			else
				'만약 슬라이드 전환속도가 슬라이드번호와 슬라이드누적카운트가 같다면 슬라이드누적카운트는 0로 하고
				'다음 이미지가 있다면 다음이미지로, 없다면 첫슬라이드로 설정한다.	
				if int(arrWG_SlideDelay(CNT1)) = int(arrSlideCNT(CNT1)) + nDelayUnitValue then 
					arrSlideCNT(CNT1) = 0
					if int(arrIDX(CNT1)) < ubound(arrWI_ImageFileName)-1 then 
						arrIDX(CNT1) = arrIDX(CNT1) + 1
					else 
						arrIDX(CNT1) = 0
					end if
				else
					arrSlideCNT(CNT1) = arrSlideCNT(CNT1) + nDelayUnitValue
				end if
			end if
			
			'해당공정에 맞는 파일중 arrIDX값에 해당하는 걸 할당 
			strImageURL(CNT1) = DefaultPath_workguide_img & arrPRD_PartNo(CNT1) & "\" & WI_ProcessName & "\" & arrWI_ImageFileName(arrIDX(CNT1))
			
		end if
	end if
next

set RS1 = nothing

strPrePartNo	= ""
strIDX 			= ""
strSlideCNT		= ""
for CNT1 = 0 to 14
	strPrePartNo	= strPrePartNo	& arrPRD_PartNo(CNT1)	&","
	strIDX			= strIDX		& arrIDX(CNT1)			&","
	strSlideCNT		= strSlideCNT	& arrSlideCNT(CNT1)		&","
next
for CNT1 = 0 to 14
	if strImageURL(CNT1) <> "" then
		
		set myImg = loadpicture(strImageURL(CNT1))
		iWidth(CNT1) = round(myImg.width / 26.4583)
		iheight(CNT1) = round(myImg.height / 26.4583)
		set myImg = nothing
		strImageURL(CNT1) = replace(strImageURL(CNT1),DefaultPath_workguide_img,"\workguide\workguide_img\")
		strImageURL(CNT1) = replace(strImageURL(CNT1),"\","/")
	else
		iWidth(CNT1) = 1
		iHeight(CNT1) = 1
		strImageURL(CNT1) = "/img/blank.gif"
	end if
	'response.write CNT1 &":"&strImageURL(CNT1) &"<br>"
next
%>

<script language="javascript">
<%
for CNT1 = 0 to 14
	if arrWG_Auto_YN(CNT1) = "Y" then
%>
if(typeof(parent.arrWorkGuide_VW[<%=CNT1%>])=='object')
{	
	if(typeof(parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide)=='object')
	{
		parent.arrWorkGuide_VW[<%=CNT1%>].document.title = "<%=arrPRD_PartNo(CNT1)%> (<%=CNT1+1%>)";
		parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.src = "<%=strImageURL(CNT1)%>?<%=replace(replace(replace(replace(replace(now()," ",""),"오후","PM"),"오전","AM"),"-",""),":","")%>";
		parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.height = <%=arrWG_ResY(CNT1)-78%>;
		parent.arrWorkGuide_VW[<%=CNT1%>].imgWorkGuide.width = parseInt(<%=arrWG_ResY(CNT1)-78%> * parseFloat(<%=iWidth(CNT1)%> / <%=iheight(CNT1)%>));
	}
}
<%
	end if
next
%>

function reload_handle()
{
	location.href="workguide_launcher_ifrm.asp?s_PRD_Line=<%=s_PRD_Line%>&strIDX=<%=server.urlencode(strIDX)%>&strSlideCNT=<%=server.urlencode(strSlideCNT)%>&strPrePartNo=<%=server.urlencode(strPrePartNo)%>&";
}

/*
function fRun()
{
	if(document.readyState == "complete")
	{
		location.href="workguide_launcher_ifrm.asp?s_PRD_Line=<%=s_PRD_Line%>&strIDX=<%=server.urlencode(strIDX)%>&strSlideCNT=<%=server.urlencode(strSlideCNT)%>&strPrePartNo=<%=server.urlencode(strPrePartNo)%>&";
	}
	else
	{
		setTimeout("fRun()",<%=nDelayUnitValue%>000);
	}
}

fRun();
*/
</script>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->