<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim CNT1
dim CNT2
dim RS1
dim SQL
dim ReSave_Require_YN

dim strBS_ST_Update_PNO
dim strBS_ST_Update_SQL
dim arrBS_ST_Update_SQL

dim strError

dim s_Work_Date
dim s_Line

dim strBOM_Sub_BS_D_No
dim strPSP_Count
dim strPSP_ST
dim strPSP_Desc
dim strPSP_Start
dim strPSP_End
dim strPSP_Sub_Start
dim strPSP_Sub_End

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

dim lstBOM_Sub_BS_D_No
dim lstPSP_Count
dim lstPSP_ST

set RS1 = Server.CreateObject("ADODB.RecordSet")

strPSP_Sub_YN = ", "&request("PSP_Sub_YN")&","

lstBOM_Sub_BS_D_No	= trim(request("lstBOM_Sub_BS_D_No"))
lstPSP_Count		= trim(request("lstPSP_Count"))
lstPSP_ST			= trim(request("lstPSP_ST"))

strBOM_Sub_BS_D_No	= replace(lstBOM_Sub_BS_D_No	,chr(13)&chr(10),",")
strPSP_Count		= replace(lstPSP_Count			,chr(13)&chr(10),",")
strPSP_ST			= replace(lstPSP_ST				,chr(13)&chr(10),",")

s_Work_Date			= request("s_Work_Date")
s_Line				= request("s_Line")

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

arrBOM_Sub_BS_D_No	= split(strBOM_Sub_BS_D_No,",")
arrPSP_Count		= split(strPSP_Count,",")
arrPSP_ST			= split(strPSP_ST,",")

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

SQL = "delete tbProcess_State_Plan where PSP_Work_Date = '"&s_Work_Date&"' and PSP_Line = '"&s_Line&"'"
sys_DBCon.execute(SQL)


strBS_ST_Update_SQL = ""
strBS_ST_Update_PNO = "-"
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
	if trim(arrBOM_Sub_BS_D_No(CNT1)) <> "" then
		arrBOM_Sub_BS_D_No(CNT1)	= trim(arrBOM_Sub_BS_D_No(CNT1))
		arrPSP_ST(CNT1)				= trim(arrPSP_ST(CNT1))
		if not(isnumeric(arrPSP_ST(CNT1))) then
			arrPSP_ST(CNT1) = 0
		end if
		'1 3 5 1 3 7 8 10 7 8 10 
		if arrBOM_Sub_BS_D_No(CNT1) <> "" and isnumeric(arrPSP_Count(CNT1)) then				
			'현재 PNO의 ST정보 가져오기
			SQL = "select BS_ST, BS_ST_ASM from tbBOM_Sub where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if RS1.Eof or RS1.Bof then '일치하는 파트넘버 없으면 패스
				
			elseif instr(strBS_ST_Update_PNO,"-"&arrBOM_Sub_BS_D_No(CNT1)&"-") = 0 then '한번 이상 SQL문에 포함된 PNO라면 생략'
				
				if isnull(RS1("BS_ST")) or isnull(RS1("BS_ST_ASM")) then 'DB상의 ST정보가 null이면...)
					
					strBS_ST_Update_SQL = strBS_ST_Update_SQL & "update tbBOM_Sub set BS_ST = "&arrPSP_ST(CNT1)&", BS_ST_ASM = "&arrPSP_ST(CNT1)&" where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'-----"
					strBS_ST_Update_PNO = strBS_ST_Update_PNO & arrBOM_Sub_BS_D_No(CNT1) & "-"
				elseif arrPSP_ST(CNT1) > 0 and (int(RS1("BS_ST")) <> int(arrPSP_ST(CNT1)) or int(RS1("BS_ST_ASM")) <> int(arrPSP_ST(CNT1))) then 'DB상의 정보와 상이하다면'
					
					strBS_ST_Update_SQL = strBS_ST_Update_SQL & "update tbBOM_Sub set BS_ST = "&arrPSP_ST(CNT1)&", BS_ST_ASM = "&arrPSP_ST(CNT1)&" where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'-----"
					strBS_ST_Update_PNO = strBS_ST_Update_PNO & arrBOM_Sub_BS_D_No(CNT1) & "-"
				end if
			end if
			RS1.Close
		end if
	end if
next
arrBS_ST_Update_SQL = split(strBS_ST_Update_SQL,"-----")
for CNT1=0 to ubound(arrBS_ST_Update_SQL)-1
	sys_DBCon.execute(arrBS_ST_Update_SQL(CNT1))
next


ReSave_Require_YN = "N"
for CNT1 = 0 to ubound(arrBOM_Sub_BS_D_No)
	if trim(arrBOM_Sub_BS_D_No(CNT1)) <> "" then
		arrBOM_Sub_BS_D_No(CNT1)	= trim(arrBOM_Sub_BS_D_No(CNT1))
		arrPSP_Count(CNT1)			= trim(arrPSP_Count(CNT1))
		arrPSP_ST(CNT1)				= trim(arrPSP_ST(CNT1))
		arrPSP_Desc(CNT1)			= trim(arrPSP_Desc(CNT1))
		arrPSP_Start(CNT1)			= trim(arrPSP_Start(CNT1))
		arrPSP_End(CNT1)			= trim(arrPSP_End(CNT1))
		arrPSP_Sub_Start(CNT1)		= trim(arrPSP_Sub_Start(CNT1))
		arrPSP_Sub_End(CNT1)		= trim(arrPSP_Sub_End(CNT1))
		if not(isnumeric(arrPSP_ST(CNT1))) then
			
			arrPSP_ST(CNT1) = 0
		end if
		
		if arrBOM_Sub_BS_D_No(CNT1) <> "" and trim(arrPSP_Start(CNT1)) = "" then
			ReSave_Require_YN = "Y"
		end if 
		
		if arrBOM_Sub_BS_D_No(CNT1) <> "" and isnumeric(arrPSP_Count(CNT1)) then
			
			if arrPSP_Count(CNT1) > 0 then
				
				SQL = "select BS_ST, BS_ST_ASM from tbBOM_Sub where BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"'"
				RS1.Open SQL,sys_DBCon
				if RS1.Eof or RS1.Bof then '일치하는 파트넘버 없으면 패스
					
					arrPSP_ST(CNT1) = 0
				else
					
					arrPSP_ST(CNT1) = RS1("BS_ST")
				end if
				RS1.Close
				
				PSP_Sub_YN = ""
				if instr(strPSP_Sub_YN,", "&cstr(CNT1)&",") > 0 then
					PSP_Sub_YN = "Y"
				end if
				
				SQL = "insert into tbProcess_State_Plan (BOM_Sub_BS_D_No, PSP_Count, PSP_ST, PSP_Desc, PSP_Start, PSP_End, PSP_Sub_YN, PSP_Sub_Start, PSP_Sub_End, PSP_Work_Date, PSP_Line) values "
				SQL = SQL & "('"&arrBOM_Sub_BS_D_No(CNT1)&"',"&arrPSP_Count(CNT1)&","&arrPSP_ST(CNT1)&",'"&arrPSP_Desc(CNT1)&"','"&arrPSP_Start(CNT1)&"','"&arrPSP_End(CNT1)&"','"&PSP_Sub_YN&"','"&arrPSP_Sub_Start(CNT1)&"','"&arrPSP_Sub_End(CNT1)&"','"&s_Work_Date&"','"&s_Line&"')"		
				sys_DBCon.execute(SQL)
				
			end if
		end if
	end if
next
set RS1 = nothing

if strError = "" then
%>
<form name="frmRedirect" action="Process_State_Plan.asp" method=post>
<input type="hidden" name="s_Work_Date"			value="<%=request("s_Work_Date")%>">
<input type="hidden" name="s_Line"				value="<%=request("s_Line")%>">
<input type="hidden" name="s_ReSave_Require_YN"	value="<%=ReSave_Require_YN%>">
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
</form>
<script language="javascript">
alert("<%=strError%>");
//parent.ifrmRecord.location.reload();
frmRedirect.submit();
</script>
<%
end if
%>


<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->