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
dim strError_temp

dim arrID_All
dim arrBOM_Sub_BS_D_No
dim arrLPE_Type
dim arrLPE_Due_Date
dim arrLPE_Req_Qty
dim arrLPE_Buyer

dim BOM_B_D_No
dim BOM_Model_BM_D_Sub_No

arrID_All				= split(Request("strID_All"),", ")
arrBOM_Sub_BS_D_No		= split(Request("BOM_Sub_BS_D_No"),", ")
arrLPE_Type				= split(Request("LPE_Type"),", ")
arrLPE_Due_Date			= split(Request("LPE_Due_Date"),", ")
arrLPE_Req_Qty			= split(Request("LPE_Req_Qty"),", ")
arrLPE_Buyer			= split(Request("LPE_Buyer"),", ")

response.write Request("LPE_Req_Qty") & "/"
response.write Request("LPE_Complete_Qty") & "/"

set RS1 = Server.CreateObject("ADODB.RecordSet")
for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrBOM_Sub_BS_D_No(CNT1)	= ucase(trim(arrBOM_Sub_BS_D_No(CNT1)))
	arrLPE_Type(CNT1)			= trim(arrLPE_Type(CNT1))
	arrLPE_Due_Date(CNT1)		= trim(arrLPE_Due_Date(CNT1))
	arrLPE_Req_Qty(CNT1)		= trim(arrLPE_Req_Qty(CNT1))
	arrLPE_Buyer(CNT1)			= trim(arrLPE_Buyer(CNT1))
next

rem DB 업데이트
for CNT1 = 0 to ubound(arrID_All)
	strError_temp = ""
			
	if strError_temp = "" then
		SQL = 		"update tbLGE_Plan_ETC set "
		SQL = SQL & "BOM_Sub_BS_D_No='"&arrBOM_Sub_BS_D_No&"', "
		SQL = SQL & "LPE_Type='"&arrLPE_Type(CNT1)&"', "
		SQL = SQL & "LPE_Due_Date='"&arrLPE_Due_Date(CNT1)&"', "
		SQL = SQL & "LPE_Req_Qty="&arrLPE_Req_Qty(CNT1)&", "
		SQL = SQL & "LPE_Buyer='"&arrLPE_Buyer(CNT1)&"' where LPE_Code='"&arrID_All(CNT1)&"'"
		sys_DBCon.execute(SQL)
	end if
	strError = strError & strError_temp
next

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
<form name="frmRedirect" action="lpe_list.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
	'strError = strError & "* 일부 항목의 수정이 취소되었습니다."
%>
<form name="frmRedirect" action="lpe_list.asp" method=post>
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