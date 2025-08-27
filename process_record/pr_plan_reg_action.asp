<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->

<%
rem 변수선언
dim SQL
dim RS1
dim CNT1

dim PR_Work_Date
dim PR_Work_Order
dim PR_Line
dim PR_Worker_CNT
dim PR_Supporter_CNT
dim BOM_Sub_BS_D_No
dim PR_Plan_Amount
dim PR_Plan_Start_Time
dim PR_Plan_End_Time
dim PR_Memo

dim PR_Plan_Time_Diff

dim arrBasicDataRestStart
dim arrBasicDataRestDiff

dim PR_Plan_Temp_Start_Time
dim PR_Plan_Temp_End_Time
dim PR_Plan_Temp_Amount

dim Time_To_Point

dim B_Code
dim BM_Code

dim pr_date

dim Flag_YN

dim temp
dim strError
dim URL_Prev
dim URL_Next

dim PR_ST
dim PR_Point

Dim PR_MPH
Dim B_IMD_MPH
Dim B_SMD_MPH
Dim B_MAN_MPH

PR_Work_Date		= trim(Request("PR_Work_Date"))
PR_Work_Order		= trim(Request("PR_Work_Order"))
PR_Line				= trim(Request("PR_Line"))
PR_Worker_CNT		= trim(Request("PR_Worker_CNT"))
PR_Supporter_CNT	= trim(Request("PR_Supporter_CNT"))
BOM_Sub_BS_D_No		= ucase(trim(Request("BOM_Sub_BS_D_No")))
PR_Plan_Amount		= trim(Request("PR_Plan_Amount"))
PR_Plan_Start_Time	= trim(Request("PR_Plan_Start_Time"))
PR_Plan_End_Time	= trim(Request("PR_Plan_End_Time"))
PR_Memo				= trim(Request("PR_Memo"))

URL_Prev			= Request("URL_Prev")
URL_Next			= Request("URL_Next")

arrBasicDataRestStart	= split(BasicDataRestStart,"-")
arrBasicDataRestDiff	= split(BasicDataRestDiff,"-")

dim LP_Model
dim strBOM_Sub_BS_D_No
dim arrBOM_Sub_BS_D_No

'PCB모델의 경우 수삽까지가 최종공정
'Assy모델의 경우 조립만 등록가능
'if BOM_Sub_BS_D_No <> "" then
'	if left(BOM_Sub_BS_D_No,4) = "6871" or left(BOM_Sub_BS_D_No,3) = "EBR" then
'		if Request("s_PR_Process") = "ASM" then
'			strError = "*해당파트넘버는 수삽(MAN)까지만 계획 입력이 가능합니다.\n"
'		end if
'	else
'		if instr("-IMD-SMD-MAN-",Request("s_PR_Process")) > 0 then
'			strError = "*해당파트넘버는 조립(ASM)만 계획 입력이 가능합니다.\n"
'		end if
'	end if
'end if

if PR_Plan_End_Time <> "" then
	if (left(PR_Plan_Start_Time,2) * 60 + right(PR_Plan_Start_Time,2)) >= (left(PR_Plan_End_Time,2) * 60 + right(PR_Plan_End_Time,2)) then
		strError = "*계획종료시각은 시작시각 이후이어야 합니다.\n"
	end if
end if

'에러메세지가 있을 경우 실행안됨
if strError = "" then
	
	PR_Plan_Start_Time	= left(PR_Plan_Start_Time,2) * 60 + right(PR_Plan_Start_Time,2)
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = "select * from tbBOM where B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No = '"&BOM_Sub_BS_D_No&"')"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		PR_ST		= 3
		B_IMD_MPH	= 180
		B_SMD_MPH	= 180
		B_MAN_MPH	= 180
	else
		PR_ST		= RS1("B_ST")
		B_IMD_MPH	= RS1("B_IMD_MPH")
		B_SMD_MPH	= RS1("B_SMD_MPH")
		B_MAN_MPH	= RS1("B_MAN_MPH")
	end if
	RS1.Close
	
	if Request("s_PR_Process") = "IMD" then
		SQL = "select sum(BQ_Qty) from tbBOM_Qty where BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' and Parts_P_P_No in (select P_P_No from tbParts where P_Work_Type in ('IMD','I/M'))"
		RS1.Open SQL,sys_DBCon
		PR_Point = RS1(0)
		RS1.Close
	elseif Request("s_PR_Process") = "SMD" then
		SQL = "select sum(BQ_Qty) from tbBOM_Qty where BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' and Parts_P_P_No in (select P_P_No from tbParts where P_Work_Type = 'SMD')"
		RS1.Open SQL,sys_DBCon
		PR_Point = RS1(0)
		RS1.Close
	else
		PR_Point = 20
	end if
	set RS1 = nothing
	
	select case PR_Line
	case "Y1" 
		Time_To_Point = Time_To_Point_Y1
	case "Y2" 
		Time_To_Point = Time_To_Point_Y2
	case "Y3" 
		Time_To_Point = Time_To_Point_Y3
	case "RH_U" 
		Time_To_Point = Time_To_Point_RH_U
	case "RHSG" 
		Time_To_Point = Time_To_Point_RHSG
	case "RH_5" 
		Time_To_Point = Time_To_Point_RH_5
	case "RHAV" 
		Time_To_Point = Time_To_Point_RHAV
	end select
	
'	if PR_Plan_End_Time = "" then
'		if instr("-IMD-SMD-",Request("s_PR_Process")) > 0 then
'			PR_Plan_Time_Diff = int(PR_Point * PR_Plan_Amount * Time_To_Point)
'		elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
'			PR_Plan_Time_Diff = int(PR_ST * PR_Plan_Amount / PR_Worker_CNT)
'		end if
'		PR_Plan_End_Time = int(PR_Plan_Start_Time) + int(PR_Plan_Time_Diff)				
'	elseif PR_Plan_End_Time <> "" then
'		PR_Plan_End_Time	= left(PR_Plan_End_Time,2) * 60 + right(PR_Plan_End_Time,2)
'		PR_Plan_Time_Diff	= PR_Plan_End_Time - PR_Plan_Start_Time
'		if instr("-IMD-SMD-",Request("s_PR_Process")) > 0 then
'			PR_Plan_Amount = int(PR_Plan_Time_Diff / Time_To_Point / PR_Point)
'		elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
'			PR_Plan_Amount = int(PR_Plan_Time_Diff * PR_Worker_CNT / PR_ST)
'		end if
'	end If

	if PR_Plan_End_Time = "" then
		if instr("-IMD-",Request("s_PR_Process")) > 0 then
			PR_Plan_Time_Diff = int(PR_Plan_Amount * 60 / B_IMD_MPH)
		elseif instr("-SMD-",Request("s_PR_Process")) > 0 then
			PR_Plan_Time_Diff = int(PR_Plan_Amount * 60 / B_SMD_MPH)
		elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
			PR_Plan_Time_Diff = int(PR_Plan_Amount * 60 / B_MAN_MPH)
		end if
		PR_Plan_End_Time = int(PR_Plan_Start_Time) + int(PR_Plan_Time_Diff)
	elseif PR_Plan_End_Time <> "" then
		PR_Plan_End_Time	= left(PR_Plan_End_Time,2) * 60 + right(PR_Plan_End_Time,2)
		PR_Plan_Time_Diff	= PR_Plan_End_Time - PR_Plan_Start_Time
		if instr("-IMD-",Request("s_PR_Process")) > 0 then
			PR_Plan_Amount = int(PR_Plan_Time_Diff * B_IMD_MPH / 60)
		elseif instr("-SMD-",Request("s_PR_Process")) > 0 then
			PR_Plan_Amount = int(PR_Plan_Time_Diff * B_SMD_MPH / 60)
		elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
			PR_Plan_Amount = int(PR_Plan_Time_Diff * B_MAN_MPH / 60)
		end if
	end If
	
	if instr("-IMD-",Request("s_PR_Process")) > 0 then
		PR_MPH = B_IMD_MPH
	elseif instr("-SMD-",Request("s_PR_Process")) > 0 then
		PR_MPH = B_SMD_MPH
	elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
		PR_MPH = B_MAN_MPH
	end if
			

	'실적 데이터 등록
	SQL = "insert into tbProcess_Record (PR_Work_Order, PR_WorkType,BOM_Sub_BS_D_No,PR_Process,PR_Worker_CNT,PR_Supporter_CNT,PR_Line,PR_Plan_Amount,PR_Work_Date,PR_Plan_Start_Time,PR_Plan_End_Time,PR_Memo,PR_ST,PR_Point,PR_MPH) values "
	SQL = SQL & "('"&PR_Work_Order&"','작업','"&BOM_Sub_BS_D_No&"','"&Request("s_PR_Process")&"',"&PR_Worker_CNT&","&PR_Supporter_CNT&",'"&PR_Line&"','"&PR_Plan_Amount&"','"&PR_Work_Date&"','"&PR_Plan_Start_Time-500&"','"&PR_Plan_End_Time-500&"','"&PR_Memo&"',"&PR_ST&","&PR_Point&","&PR_MPH&")"
	sys_DBCon.execute(SQL)
	
	'if PR_Work_Order <> "" and int(PR_Plan_Amount) > 0 and PR_Plan_Start_Time < PR_Plan_End_Time then
		'SQL = "update tbLGE_Plan set LP_"&Request("s_PR_Process")&"_Complete_YN = '○' where LP_Work_Order='"&PR_Work_Order&"'"
		'sys_DBCon.execute(SQL)
	'end if
end if
%>

<%
dim Request_Fields
dim strRequestForm
dim strRequestQueryString
for each Request_Fields in Request.Form
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
for each Request_Fields in Request.QueryString
	if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	end if
next
if strError = "" then
%>
<form name="frmRedirect" action="pr_plan_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="pr_plan_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>

<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->