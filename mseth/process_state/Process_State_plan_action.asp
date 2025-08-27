<!-- #include Virtual = "/mseth/header/asp_header.asp" -->
<!-- include Virtual = "/mseth/header/session_check_header.asp" -->
<!-- #include Virtual = "/mseth/header/db_header.asp" -->
<!-- #include Virtual = "/mseth/header/inc_share_function.asp" -->
<%
'변수 선언
dim CNT1
dim CNT2
dim RS1
dim SQL

'재저장 필요 플래그
dim ReSave_Require_YN

'SQL문자열을 
dim strBS_ST_Update_PNO
dim strBS_ST_Update_SQL
dim arrBS_ST_Update_SQL

'에러 문자열
dim strError

'계획 날짜, 라인 변수
dim s_Work_Date
dim s_Line

'리스트용 문자열 배열
dim strBOM_Sub_BS_D_No
dim strPSP_Count
dim strPSP_ST
dim strPSP_Desc
dim strPSP_Start
dim strPSP_End
dim strPSP_Sub_Start
dim strPSP_Sub_End

'리스트용 배열
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

'textarea로 값을 받는 경우
dim lstBOM_Sub_BS_D_No
dim lstPSP_Count
dim lstPSP_ST

set RS1 = Server.CreateObject("ADODB.RecordSet")

strPSP_Sub_YN = ", "&request("PSP_Sub_YN")&","

lstBOM_Sub_BS_D_No	= trim(request("lstBOM_Sub_BS_D_No"))
lstPSP_Count		= trim(request("lstPSP_Count"))
lstPSP_ST			= trim(request("lstPSP_ST"))

'textarea 변수값을 받아서, 리스트 변수값화 시킴
strBOM_Sub_BS_D_No	= replace(lstBOM_Sub_BS_D_No	,chr(13)&chr(10),",")
strPSP_Count		= replace(lstPSP_Count			,chr(13)&chr(10),",")
strPSP_ST			= replace(lstPSP_ST				,chr(13)&chr(10),",")

s_Work_Date			= request("s_Work_Date")
s_Line				= request("s_Line")

'textarea 변수값이 없다면 > 즉 리스트 변수값이라면
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

'우선 파트넘버, 수량, st만 배열화
arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,",")
arrPSP_Count		= split(strPSP_Count,",")
arrPSP_ST			= split(strPSP_ST,",")

'textarea에 st를 빼먹었다면, 빈 배열은 잡아줌.
if trim(lstBOM_Sub_BS_D_No) <> "" and trim(lstPSP_ST) = "" then
	redim arrPSP_ST(ubound(arrBOM_Sub_BS_D_No))
end if

'textarea로 받은 경우라면, 빈 배열 잡아주고, 리스트로 받은 경우 배열화
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

'기존 계획 정보를 삭제한다.
SQL = "delete tbProcess_State_Plan where PSP_Work_Date = '"&s_Work_Date&"' and PSP_Line = '"&s_Line&"'"
sys_DBCon.execute(SQL)

'st가 db와 다른 경우가 있다면 업데이트를 해야 하기 때문에, 로그 문자열을 사용한다. 우선 초기화
strBS_ST_Update_SQL = ""
strBS_ST_Update_PNO = "-"

'배열을 루핑
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
  '파트넘버가 유효하다면
	if trim(arrBOM_Sub_BS_D_No(CNT1)) <> "" and len(trim(arrBOM_Sub_BS_D_No(CNT1)))=11 and isnumeric(arrPSP_Count(CNT1)) then
		'공백 제거 및 파트넘버 대문자 변환
		arrBOM_Sub_BS_D_No(CNT1)	= ucase(trim(arrBOM_Sub_BS_D_No(CNT1)))
		arrPSP_ST(CNT1)						= trim(arrPSP_ST(CNT1))
		
				
		'db에서 해당 파트넘버의 st값을 찾는다.
		SQL = "select BS_ST, BS_ST_ASM from tbBOM_Sub where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'"
		RS1.Open SQL,sys_DBCon
		
		'일치하는 파트넘버가 없다면.
		if RS1.Eof or RS1.Bof then
			'ST값이 비어서 왔다면, 10초로 셋
			if arrPSP_ST(CNT1) = "" or not(isnumeric(arrPSP_ST(CNT1))) then
				arrPSP_ST(CNT1) = "10"
			end if
			'이 파트넘버를 디폴드 st값으로 등록한다.
			SQL = "insert into tbBOM_Sub (BS_D_No, BS_ST, BS_ST_ASM) values ('"&arrBOM_Sub_BS_D_No(CNT1)&"',"&arrPSP_ST(CNT1)&","&arrPSP_ST(CNT1)&")"
			sys_DBCon.execute(SQL)

		'일치하는 파트넘버가 있다면                                                                                                                                                                                                       
		else
			'새로 입력된 ST가 접합하면 DB정보 무시, 부적합하면, DB의 ST를 활용
			if not(isnumeric(arrPSP_ST(CNT1))) then
				arrPSP_ST(CNT1) = RS1("BS_ST")
			end if
			
			if instr(strBS_ST_Update_PNO,"-"&arrBOM_Sub_BS_D_No(CNT1)&"-") = 0 then '한번 이상 SQL문에 포함된 PNO라면 생략'
				if isnull(RS1("BS_ST")) or isnull(RS1("BS_ST_ASM")) then 'DB상의 ST정보가 null이면. 새로운 ST로 업데이트
					strBS_ST_Update_SQL = strBS_ST_Update_SQL & "update tbBOM_Sub set BS_ST = "&arrPSP_ST(CNT1)&", BS_ST_ASM = "&arrPSP_ST(CNT1)&" where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'-----"
					strBS_ST_Update_PNO = strBS_ST_Update_PNO & arrBOM_Sub_BS_D_No(CNT1) & "-"
				elseif int(RS1("BS_ST")) <> int(arrPSP_ST(CNT1)) or int(RS1("BS_ST_ASM")) <> int(arrPSP_ST(CNT1)) then 'DB상의 정보와 상이하다면, 새로운 ST로 업데이트'
					strBS_ST_Update_SQL = strBS_ST_Update_SQL & "update tbBOM_Sub set BS_ST = "&arrPSP_ST(CNT1)&", BS_ST_ASM = "&arrPSP_ST(CNT1)&" where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'-----"
					strBS_ST_Update_PNO = strBS_ST_Update_PNO & arrBOM_Sub_BS_D_No(CNT1) & "-"
				end if
			end if
		end if
		RS1.Close
	end if
next

'몰아서 업데이트
arrBS_ST_Update_SQL = split(strBS_ST_Update_SQL,"-----")
for CNT1=0 to ubound(arrBS_ST_Update_SQL)-1
	sys_DBCon.execute(arrBS_ST_Update_SQL(CNT1))
next

'우선 재저장 플래그를 N으로 둔다.
ReSave_Require_YN = "N"

'다시 루핑
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
	'파트넘버가 유효하면,
	if trim(arrBOM_Sub_BS_D_No(CNT1)) <> "" and len(trim(arrBOM_Sub_BS_D_No(CNT1)))=11 then
		'공백제거
		arrBOM_Sub_BS_D_No(CNT1)	= trim(arrBOM_Sub_BS_D_No(CNT1))
		arrPSP_Count(CNT1)			= trim(arrPSP_Count(CNT1))
		arrPSP_ST(CNT1)				= trim(arrPSP_ST(CNT1))
		arrPSP_Desc(CNT1)			= trim(arrPSP_Desc(CNT1))
		arrPSP_Start(CNT1)			= trim(arrPSP_Start(CNT1))
		arrPSP_End(CNT1)			= trim(arrPSP_End(CNT1))
		arrPSP_Sub_Start(CNT1)		= trim(arrPSP_Sub_Start(CNT1))
		arrPSP_Sub_End(CNT1)		= trim(arrPSP_Sub_End(CNT1))
		
		'ST가 숫자가 아니면 10초로 셋
		if not(isnumeric(arrPSP_ST(CNT1))) then
			arrPSP_ST(CNT1) = 10
		end if
		
		'파트넘버는 있는데, 시작시간이 없다면, 재저장 플래그 Y
		if arrBOM_Sub_BS_D_No(CNT1) <> "" and trim(arrPSP_Start(CNT1)) = "" then
			ReSave_Require_YN = "Y"
		end if 
		
		'파트넘버 유효하고, 목표수량도 유효하다면
		if arrBOM_Sub_BS_D_No(CNT1) <> "" and isnumeric(arrPSP_Count(CNT1)) then
			'목표수량이 정말 0보다 크고 유효하다면,
			if arrPSP_Count(CNT1) > 0 then
				'DB의 ST를 가져옴
				SQL = "select BS_ST, BS_ST_ASM from tbBOM_Sub where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'"
				RS1.Open SQL,sys_DBCon
				if RS1.Eof or RS1.Bof then '일치하는 파트넘버 없으면 st 10초로 지정
					arrPSP_ST(CNT1) = 10
				else
					arrPSP_ST(CNT1) = RS1("BS_ST") '일치하는 파트넘버 있으면 그 값을 가져옴.
				end if
				RS1.Close
				
				'서브PCB관련
				PSP_Sub_YN = ""
				if instr(strPSP_Sub_YN,", "&cstr(CNT1)&",") > 0 then
					PSP_Sub_YN = "Y"
				end if
				
				'계획DB에 넣음
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