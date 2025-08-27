<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim RS1
dim SQL
dim CNT1

dim strError
dim strError_Temp

dim arrID_All
dim arrPR_Work_Date
dim arrPR_Work_Order
dim arrPR_Line
dim arrPR_Worker_CNT
dim arrPR_Supporter_CNT
dim arrBOM_Sub_BS_D_No
dim arrPR_Plan_Amount
dim arrPR_Plan_Start_Time
dim arrPR_Plan_End_Time
dim arrPR_Memo

dim PR_WorkType
dim BOM_Sub_BS_D_No
dim PR_Process
dim PR_Amount

dim PR_ST
dim PR_Point
dim Time_To_Point
dim PR_Plan_Time_Diff

Dim PR_MPH
Dim B_IMD_MPH
Dim B_SMD_MPH
Dim B_MAN_MPH

arrID_All				= split(Request("strID_All")&" "			,", ")
arrPR_Work_Date			= split(Request("PR_Work_Date")&" "			,", ")
arrPR_Work_Order		= split(Request("PR_Work_Order")&" "		,", ")
arrPR_Line				= split(Request("PR_Line")&" "				,", ")
arrPR_Worker_CNT		= split(Request("PR_Worker_CNT")&" "		,", ")
arrPR_Supporter_CNT		= split(Request("PR_Supporter_CNT")&" "		,", ")
arrBOM_Sub_BS_D_No		= split(Request("BOM_Sub_BS_D_No")&" "		,", ")
arrPR_Plan_Amount		= split(Request("PR_Plan_Amount")&" "		,", ")
arrPR_Plan_Start_Time	= split(Request("PR_Plan_Start_Time")&" "	,", ")
arrPR_Plan_End_Time		= split(Request("PR_Plan_End_Time")&" "		,", ")
arrPR_Memo				= split(Request("PR_Memo")&" "				,", ")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrPR_Work_Date(CNT1)		= trim(arrPR_Work_Date(CNT1))
	arrPR_Work_Order(CNT1)		= trim(arrPR_Work_Order(CNT1))
	arrPR_Line(CNT1)			= trim(arrPR_Line(CNT1))
	arrPR_Worker_CNT(CNT1)		= trim(arrPR_Worker_CNT(CNT1))
	arrPR_Supporter_CNT(CNT1)	= trim(arrPR_Supporter_CNT(CNT1))
	arrBOM_Sub_BS_D_No(CNT1)	= trim(arrBOM_Sub_BS_D_No(CNT1))
	arrPR_Plan_Amount(CNT1)		= trim(arrPR_Plan_Amount(CNT1))
	arrPR_Plan_Start_Time(CNT1)	= trim(arrPR_Plan_Start_Time(CNT1))
	arrPR_Plan_End_Time(CNT1)	= trim(arrPR_Plan_End_Time(CNT1))
	arrPR_Memo(CNT1)			= trim(arrPR_Memo(CNT1))
next

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""	
		
		if arrPR_Plan_End_Time(CNT1) <> "" then
			if (left(arrPR_Plan_Start_Time(CNT1),2) * 60 + right(arrPR_Plan_Start_Time(CNT1),2)) >= (left(arrPR_Plan_End_Time(CNT1),2) * 60 + right(arrPR_Plan_End_Time(CNT1),2)) then
				strError_Temp = "*["&CNT1&"번째 항목] 종료시각은 시작시각 이후이어야 합니다.\n"
			end if
		end if
	
		if strError_Temp = "" then
			
			arrPR_Plan_Start_Time(CNT1)	= left(arrPR_Plan_Start_Time(CNT1),2) * 60 + right(arrPR_Plan_Start_Time(CNT1),2)
			
			SQL = "select * from tbBOM where B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"')"
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
			
			if PR_Process = "IMD" then
				SQL = "select sum(BQ_Qty) from tbBOM_Qty where BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"' and Parts_P_P_No in (select P_P_No from tbParts where P_Work_Type in ('IMD','I/M'))"
				RS1.Open SQL,sys_DBCon
				PR_Point = RS1(0)
				RS1.Close
			elseif PR_Process = "SMD" then
				SQL = "select sum(BQ_Qty) from tbBOM_Qty where BOM_Sub_BS_D_No = '"&arrBOM_Sub_BS_D_No(CNT1)&"' and Parts_P_P_No in (select P_P_No from tbParts where P_Work_Type = 'SMD')"
				RS1.Open SQL,sys_DBCon
				PR_Point = RS1(0)
				RS1.Close
			else
				PR_Point = 20
			end if
			
			select case arrPR_Line(CNT1)
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
			end Select
			
'			if arrPR_Plan_End_Time(CNT1) = "" then
'				if instr("-IMD-SMD-",Request("s_PR_Process")) > 0 then
'					PR_Plan_Time_Diff = int(PR_Point * arrPR_Plan_Amount(CNT1) * Time_To_Point)
'				elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
'					PR_Plan_Time_Diff = int(PR_ST * arrPR_Plan_Amount(CNT1) / arrPR_Worker_CNT(CNT1))
'				end if
'				arrPR_Plan_End_Time(CNT1) = int(arrPR_Plan_Start_Time(CNT1)) + int(PR_Plan_Time_Diff)
'			elseif arrPR_Plan_End_Time(CNT1) <> "" then
'				arrPR_Plan_End_Time(CNT1)	= left(arrPR_Plan_End_Time(CNT1),2) * 60 + right(arrPR_Plan_End_Time(CNT1),2)
'				PR_Plan_Time_Diff	= arrPR_Plan_End_Time(CNT1) - arrPR_Plan_Start_Time(CNT1)
'				if instr("-IMD-SMD-",Request("s_PR_Process")) > 0 then
'					arrPR_Plan_Amount(CNT1) = int(PR_Plan_Time_Diff / Time_To_Point / PR_Point)
'				elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
'					arrPR_Plan_Amount(CNT1) = int(PR_Plan_Time_Diff * arrPR_Worker_CNT(CNT1) / PR_ST)
'				end if
'			end If
			
			if arrPR_Plan_End_Time(CNT1) = "" then
				if instr("-IMD-",Request("s_PR_Process")) > 0 then
					PR_Plan_Time_Diff = int(arrPR_Plan_Amount(CNT1) * 60 / B_IMD_MPH)
				elseif instr("-SMD-",Request("s_PR_Process")) > 0 then
					PR_Plan_Time_Diff = int(arrPR_Plan_Amount(CNT1) * 60 / B_SMD_MPH)
				elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
					PR_Plan_Time_Diff = int(arrPR_Plan_Amount(CNT1) * 60 / B_MAN_MPH)
				end if
				arrPR_Plan_End_Time(CNT1) = int(arrPR_Plan_Start_Time(CNT1)) + int(PR_Plan_Time_Diff)
			elseif arrPR_Plan_End_Time(CNT1) <> "" then
				arrPR_Plan_End_Time(CNT1)	= left(arrPR_Plan_End_Time(CNT1),2) * 60 + right(arrPR_Plan_End_Time(CNT1),2)
				PR_Plan_Time_Diff	= arrPR_Plan_End_Time(CNT1) - arrPR_Plan_Start_Time(CNT1)
				if instr("-IMD-",Request("s_PR_Process")) > 0 then
					arrPR_Plan_Amount(CNT1) = int(PR_Plan_Time_Diff * B_IMD_MPH / 60)
				elseif instr("-SMD-",Request("s_PR_Process")) > 0 then
					arrPR_Plan_Amount(CNT1) = int(PR_Plan_Time_Diff * B_SMD_MPH / 60)
				elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
					arrPR_Plan_Amount(CNT1) = int(PR_Plan_Time_Diff * B_MAN_MPH / 60)
				end if
			end If
			
			if instr("-IMD-",Request("s_PR_Process")) > 0 then
				PR_MPH = B_IMD_MPH
			elseif instr("-SMD-",Request("s_PR_Process")) > 0 then
				PR_MPH = B_SMD_MPH
			elseif instr("-MAN-ASM-",Request("s_PR_Process")) > 0 then
				PR_MPH = B_MAN_MPH
			end if
			
			'SQL = "select top 1 sum(PR_Amount) from tbProcess_Record where PR_Work_Order='"&arrPR_Work_Order(CNT1)&"' and PR_Code <> '"&arrID_All(CNT1)&"' and PR_Process = '"&PR_Process&"'"
			'RS1.Open SQL,sys_DBCon
			'if RS1.Eof or RS1.Bof then
				'if arrPR_Work_Order(CNT1) <> "" then
					'SQL = "update tbLGE_Plan set LP_"&PR_Process&"_Complete_YN = '' where LP_Work_Order='"&arrPR_Work_Order(CNT1)&"'"
					'sys_DBCon.execute(SQL)
				'end if
			'else
				'if arrPR_Work_Order(CNT1) <> "" and RS1(0) > 0 then
					'SQL = "update tbLGE_Plan set LP_"&PR_Process&"_Complete_YN = 'Y' where LP_Work_Order='"&arrPR_Work_Order(CNT1)&"'"
					'sys_DBCon.execute(SQL)
				'end if
			'end if	
			'RS1.Close
			
			SQL = "update tbProcess_Record set "
			SQL = SQL & "	PR_Work_Order='"&arrPR_Work_Order(CNT1)&"', "
			SQL = SQL & "	PR_Work_Date='"&arrPR_Work_Date(CNT1)&"', "
			SQL = SQL & "	PR_Line='"&arrPR_Line(CNT1)&"', "
			SQL = SQL & "	PR_Worker_CNT="&arrPR_Worker_CNT(CNT1)&", "
			SQL = SQL & "	PR_Supporter_CNT="&arrPR_Supporter_CNT(CNT1)&", "
			SQL = SQL & "	BOM_Sub_BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"', "
			SQL = SQL & "	PR_ST="&PR_ST&", "
			SQL = SQL & "	PR_Point="&PR_Point&", "
			SQL = SQL & "	PR_MPH="&PR_MPH&", "
			SQL = SQL & "	PR_Plan_Amount="&arrPR_Plan_Amount(CNT1)&", "
			SQL = SQL & "	PR_Plan_Start_Time='"&arrPR_Plan_Start_Time(CNT1)-500&"', "
			SQL = SQL & "	PR_Plan_End_Time='"&arrPR_Plan_End_Time(CNT1)-500&"', "
			SQL = SQL & "	PR_Memo='"&arrPR_Memo(CNT1)&"' "			
			SQL = SQL & "where PR_Code='"&arrID_All(CNT1)&"' "
			sys_DBCon.execute(SQL)
			
		end if
		
		strError = strError & strError_Temp
	next
end if

set RS1 = nothing
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
	'strError = strError & "* 일부의 수정이 취소되었습니다."
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