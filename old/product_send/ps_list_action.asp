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
dim arrBOM_Sub_BS_D_No
dim arrPS_Send_Date
dim arrPS_Qty
dim arrLGE_Plan_LP_Work_Order
dim arrLGE_Plan_ETC_LPE_Code

arrID_All					= split(Request("strID_All")&" "				,", ")
arrBOM_Sub_BS_D_No			= split(Request("BOM_Sub_BS_D_No")&" "			,", ")
arrPS_Send_Date				= split(Request("PS_Send_Date")&" "				,", ")
arrPS_Qty					= split(Request("PS_Qty")&" "					,", ")
arrLGE_Plan_LP_Work_Order	= split(Request("LGE_Plan_LP_Work_Order")&" "	,", ")
arrLGE_Plan_ETC_LPE_Code	= split(Request("LGE_Plan_ETC_LPE_Code")&" "	,", ")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)					= trim(arrID_All(CNT1))
	arrBOM_Sub_BS_D_No(CNT1)		= trim(arrBOM_Sub_BS_D_No(CNT1))
	arrPS_Send_Date(CNT1)			= trim(arrPS_Send_Date(CNT1))
	arrPS_Qty(CNT1)					= trim(arrPS_Qty(CNT1))
	arrLGE_Plan_LP_Work_Order(CNT1)	= trim(arrLGE_Plan_LP_Work_Order(CNT1))
	arrLGE_Plan_ETC_LPE_Code(CNT1)	= trim(arrLGE_Plan_ETC_LPE_Code(CNT1))
next

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""		
	
		if strError_Temp = "" then
			SQL = "update tbProduct_Send set "
			SQL = SQL & "	PS_Send_Date='"&arrPS_Send_Date(CNT1)&"', "
			SQL = SQL & "	PS_Qty="&arrPS_Qty(CNT1)&", "
			SQL = SQL & "	LGE_Plan_LP_Work_Order='"&arrLGE_Plan_LP_Work_Order(CNT1)&"', "
			if isnumeric(arrLGE_Plan_ETC_LPE_Code(CNT1)) then
				SQL = SQL & "	LGE_Plan_ETC_LPE_Code="&arrLGE_Plan_ETC_LPE_Code(CNT1)&", "
			end if
			SQL = SQL & "	BOM_Sub_BS_D_No='"&arrBOM_Sub_BS_D_No(CNT1)&"' "
			SQL = SQL & "where PS_Code='"&arrID_All(CNT1)&"' "
			sys_DBCon.execute(SQL)
		end if
		
		strError = strError & strError_Temp
	next
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
<form name="frmRedirect" action="ps_list.asp" method=post>
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
<form name="frmRedirect" action="ps_list.asp" method=post>
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