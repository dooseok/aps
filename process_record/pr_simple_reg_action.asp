<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->

<%
rem 변수선언
dim SQL
dim RS1
dim CNT1

dim PR_Work_Order
dim PR_WorkType
dim PR_Process
dim PR_Line
dim BOM_Sub_BS_D_No
dim PR_Amount
dim PR_Amount_NG
dim PR_Worker_CNT
dim PR_Work_Date
dim PR_Start_Time
dim PR_End_Time
dim PR_Loss_Time
dim PR_Rest_Time
dim PR_Memo
dim PR_Point
dim PR_ST

dim B_Code
dim BM_Code

dim pr_date

dim temp
dim strError
dim URL_Prev
dim URL_Next

PR_Work_Order		= trim(Request("PR_Work_Order"))
PR_WorkType			= trim(Request("PR_WorkType"))
PR_Process			= trim(Request("PR_Process"))
PR_Line				= trim(Request("PR_Line"))
BOM_Sub_BS_D_No		= ucase(trim(Request("BOM_Sub_BS_D_No")))
PR_Amount			= trim(Request("PR_Amount"))
PR_Amount_NG		= trim(Request("PR_Amount_NG"))
PR_Worker_CNT		= trim(Request("PR_Worker_CNT"))
PR_Work_Date		= trim(Request("PR_Work_Date"))
PR_Start_Time		= trim(Request("PR_Start_Time"))
PR_End_Time			= trim(Request("PR_End_Time"))
PR_Loss_Time		= trim(Request("PR_Loss_Time"))
PR_Rest_Time		= trim(Request("PR_Rest_Time"))
PR_Memo				= trim(Request("PR_Memo"))
PR_Point			= trim(Request("PR_Point"))
PR_ST				= trim(Request("PR_ST"))

pr_date				= trim(Request("pr_date"))

URL_Prev			= Request("URL_Prev")
URL_Next			= Request("URL_Next")

dim LP_Model
dim strBOM_Sub_BS_D_No
dim arrBOM_Sub_BS_D_No

if PR_Start_Time > PR_End_Time then
	strError = "*작업시간이 잘못되었습니다.\n"
end if

'에러메세지가 있을 경우 실행안됨

set RS1 = Server.CreateObject("ADODB.RecordSet")

if strError = "" then
	'실적 데이터 등록
	if PR_WorkType <> "작업" then
		PR_Work_Order = ""
	end if
	
	PR_Start_Time	= left(PR_Start_Time,2) * 60 + right(PR_Start_Time,2) - 500
	PR_End_Time		= left(PR_End_Time,2) * 60 + right(PR_End_Time,2) - 500
	
	SQL = "select * from tbBOM where B_Code in (select BOM_B_Code from tbBOM_Sub where BS_D_No = '"&BOM_Sub_BS_D_No&"')"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		PR_ST		= 3
	else
		if Request("s_PR_Process") = "MAN" then
			PR_ST		= RS1("B_ST")
		else
			PR_ST		= RS1("B_ST_Assm")
		end if
	end if
	RS1.Close

	if Request("s_PR_Process") = "IMD" then
		if instr(lcase(PR_Line),"rh") > 0 then
			if isnumeric(left(BOM_Sub_BS_D_No,3)) then
				SQL = "select avg(bs_imd_radial_point) from tbBOM_Sub where left(BS_D_No,10) = '"&BOM_Sub_BS_D_No&"'"
			else
				SQL = "select avg(bs_imd_radial_point) from tbBOM_Sub where left(BS_D_No,9) = '"&BOM_Sub_BS_D_No&"'"
			end if
		else
			if isnumeric(left(BOM_Sub_BS_D_No,3)) then
				SQL = "select avg(bs_imd_axial_point) from tbBOM_Sub where left(BS_D_No,10) = '"&BOM_Sub_BS_D_No&"'"
			else
				SQL = "select avg(bs_imd_axial_point) from tbBOM_Sub where left(BS_D_No,9) = '"&BOM_Sub_BS_D_No&"'"
			end if
		end if
		RS1.Open SQL,sys_DBCon
		PR_Point = RS1(0)
		RS1.Close

	elseif Request("s_PR_Process") = "SMD" then
		SQL = "select sum(BQ_Qty) from tbBOM_Qty where BOM_Sub_BS_D_No = '"&BOM_Sub_BS_D_No&"' and Parts_P_P_No in (select P_P_No from tbParts where P_Work_Type = 'SMD')"
		RS1.Open SQL,sys_DBCon
		PR_Point = RS1(0)
		RS1.Close
	end if

	if isnumeric(PR_Point) and PR_Point <> "" and not(ISNULL(PR_Point)) then
	else
		PR_Point = 20
	end if

	SQL = "insert into tbProcess_Record (PR_Work_Order, PR_WorkType,BOM_Sub_BS_D_No,PR_Process,PR_Worker_CNT,PR_Line,PR_Amount,PR_Amount_NG,PR_Work_Date,PR_Start_Time,PR_End_Time,PR_Loss_Time,PR_Rest_Time,PR_Plan_Start_Time,PR_Plan_End_Time,PR_Memo,PR_Point,PR_ST) values "
	SQL = SQL & "('"&PR_Work_Order&"','"&PR_WorkType&"','"&BOM_Sub_BS_D_No&"','"&Request("s_PR_Process")&"',"&PR_Worker_CNT&",'"&PR_Line&"','"&PR_Amount&"','"&PR_Amount_NG&"','"&PR_Work_Date&"','"&PR_Start_Time&"','"&PR_End_Time&"',"&PR_Loss_Time&","&PR_Rest_Time&",'','','"&PR_Memo&"',"&PR_Point&","&PR_ST&")"
	sys_DBCon.execute(SQL)

	SQL = "select BOM_Sub_BS_D_No from tbBOM_Qty where BQ_Qty > 0 and Parts_P_P_No = '"&BOM_Sub_BS_D_No&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if Request("s_PR_Process") = "MAN" and PR_WorkType = "작업" then
			'SQL = "insert into tbProcess_Record (PR_Work_Order, PR_WorkType,BOM_Sub_BS_D_No,PR_Process,PR_Worker_CNT,PR_Line,PR_Amount,PR_Amount_NG,PR_Work_Date,PR_Start_Time,PR_End_Time,PR_Loss_Time,PR_Rest_Time,PR_Plan_Start_Time,PR_Plan_End_Time,PR_Memo,PR_Point,PR_ST) values "
			'SQL = SQL & "('"&PR_Work_Order&"','"&PR_WorkType&"','"&BOM_Sub_BS_D_No&"','ASM',1,'C1','"&PR_Amount&"',0,'"&PR_Work_Date&"','"&PR_Start_Time&"','"&PR_End_Time&"',"&PR_Loss_Time&","&PR_Rest_Time&",'','','"&PR_Memo&"',"&PR_Point&","&PR_ST&")"
			'sys_DBCon.execute(SQL)
		end if
	end if
	RS1.Close

	if PR_WorkType = "작업" and PR_Amount > 0 and instr(PR_Line,"RH") > 0  then
		'입력된 실적에 해당하는 모델파트넘버의 해당공정 재고를 +시킴
		if Request("s_PR_Process") <> "DLV" then
			call Process_Qty_BOM_Sub_Plus(BOM_Sub_BS_D_No,Request("s_PR_Process"),PR_Amount, PR_Work_Date)
		end if
		
		'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 -시킴
		call Process_Qty_BOM_Sub_Before_Minus(BOM_Sub_BS_D_No,Request("s_PR_Process"),PR_Amount, PR_Work_Date)

		'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 -시킴
		if Request("s_PR_Process") <> "DLV" then
			call Process_Qty_Parts_Minus(BOM_Sub_BS_D_No,Request("s_PR_Process"),PR_Amount)
		end if
	end if
	
	dim arrTemp
	SQL = "select sum(PR_Amount) from tbProcess_Record where PR_Work_Order='"&PR_Work_Order&"' and PR_Process='"&Request("s_PR_Process")&"'"
	RS1.Open SQL,sys_DBCon
	if instr(PR_Work_Order,"_") > 0 then
		arrTemp = split(PR_Work_Order,"_")
		SQL = "update tbLGE_Plan_ETC set LPE_"&Request("s_PR_Process")&"_Complete_Qty = "&RS1(0)&" where LPE_Type='"&arrTemp(1)&"' and LPE_Code='"&arrTemp(0)&"'"
	else
		SQL = "update tbLGE_Plan set LP_"&Request("s_PR_Process")&"_Complete_Qty = "&RS1(0)&" where LP_Work_Order='"&PR_Work_Order&"'"
	end if
	'sys_DBCon.execute(SQL)
	RS1.Close
	
	'if PR_Work_Order <> "" and int(PR_Amount) > 0 and PR_Start_Time < PR_End_Time then
		'SQL = "update tbLGE_Plan set LP_"&Request("s_PR_Process")&"_Complete_YN = 'Y' where LP_Work_Order='"&PR_Work_Order&"'"
		'sys_DBCon.execute(SQL)
	'end if
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
<form name="frmRedirect" action="pr_simple_list.asp" method=post>

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
<form name="frmRedirect" action="pr_simple_list.asp" method=post>

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