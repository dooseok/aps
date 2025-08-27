<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim SQL
dim RS1

dim PR_Code

dim PR_Work_Order
dim PR_WorkType
dim BOM_Sub_BS_D_No
dim PR_Process
dim PR_Amount


dim strError

PR_Code = request("PR_Code")

set RS1 = Server.CreateObject("ADODB.RecordSet")

'삭제하기 전 내용을 조회
SQL = "select top 1 PR_Work_Order,PR_WorkType,BOM_Sub_BS_D_No,PR_Process,PR_Amount from tbProcess_Record where PR_Code = '"&PR_Code&"'"
RS1.Open SQL,sys_DBCon
PR_Work_Order	= RS1("PR_Work_Order")
PR_WorkType		= RS1("PR_WorkType")
BOM_Sub_BS_D_No	= RS1("BOM_Sub_BS_D_No")
PR_Process		= RS1("PR_Process")
PR_Amount		= RS1("PR_Amount")
RS1.Close

if PR_WorkType = "작업" t and PR_Amount > 0 then
	if PR_Process <> "DLV" then
		'입력된 실적에 해당하는 모델 파트넘버의 해당공정 재고를 -시킴
		call Process_Qty_BOM_Sub_Minus(BOM_Sub_BS_D_No,PR_Process,PR_Amount)
	end if
	
	'입력된 실적에 해당하는 모델파트넘버의 이전공정 재고를 +시킴
	call Process_Qty_BOM_Sub_Before_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount)
	
	if PR_Process <> "DLV" then
		'입력된 실적에 해당하는 모델파트넘버의 해당공정에서 쓰이는 자재재고를 +시킴
		call Process_Qty_Parts_Plus(BOM_Sub_BS_D_No,PR_Process,PR_Amount)
	end if
	
	'SQL = "select top 1 PR_Code from tbProcess_Record where PR_Work_Order='"&PR_Work_Order&"' and PR_Code <> '"&PR_Code&"' and PR_Process = '"&PR_Process&"'"
	'RS1.Open SQL,sys_DBCon
	'if RS1.Eof or RS1.Bof then
		'if PR_Work_Order <> "" then
			'SQL = "update tbLGE_Plan set LP_"&Request("s_PR_Process")&"_Complete_YN = '' where LP_Work_Order='"&PR_Work_Order&"'"
			'sys_DBCon.execute(SQL)
		'end if
	'else
		'if PR_Work_Order <> "" then
			'SQL = "update tbLGE_Plan set LP_"&Request("s_PR_Process")&"_Complete_YN = 'Y' where LP_Work_Order='"&PR_Work_Order&"'"
			'sys_DBCon.execute(SQL)
		'end if	
	'end if	
	'RS1.Close
	
	dim arrTemp
	SQL = "select sum(PR_Amount) from tbLGE_Plan where PR_Work_Order='"&PR_Work_Order&"' and PR_Process='"&PR_Process&"'"
	RS1.Open SQL,sys_DBCon
	if instr(PR_Work_Order,"_") > 0 then
		arrTemp = split(PR_Work_Order,"_")
		SQL = "update tbLGE_Plan_ETC set LPE_"&PR_Process&"_Complete_Qty = "&RS1(0)&" where LPE_Type='"&arrTemp(1)&"' and LPE_Code='"&arrTemp(0)&"'"
	else
		SQL = "update tbLGE_Plan set LP_"&PR_Process&"_Complete_Qty = "&RS1(0)&" where LP_Work_Order='"&PR_Work_Order&"'"
	end if
	sys_DBCon.execute(SQL)
	RS1.Close

end if
set RS1 = nothing


'실적 정보 삭제
SQL = "delete from tbProcess_Record where PR_Code='"&PR_Code&"'"
sys_DBCon.execute(SQL)
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
<form name="frmRedirect" action="pr_list.asp" method=post>

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
<form name="frmRedirect" action="pr_list.asp" method=post>

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