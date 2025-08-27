<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
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
dim arrBPL_Market
dim arrBPL_Currency
dim arrBPL_Price
dim arrPartner_P_Name
dim arrBPL_Apply_Date

arrID_All			= split(Request("strID_All")&" "	,", ")
arrBPL_Market		= split(Request("BPL_Market")&" "	,", ")
arrBPL_Currency		= split(Request("BPL_Currency")&" "	,", ")
arrBPL_Price		= split(Request("BPL_Price")&" "	,", ")
arrPartner_P_Name		= split(Request("Partner_P_Name")&" "	,", ")
arrBPL_Apply_Date	= split(Request("BPL_Apply_Date")&" "	,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		if strError_Temp = "" then
			SQL = "update tbBOM_Price_Log set "
			SQL = SQL & "	BPL_Market='"&trim(arrBPL_Market(CNT1))&"', "
			SQL = SQL & "	BPL_Currency='"&trim(arrBPL_Currency(CNT1))&"', "
			SQL = SQL & "	BPL_Price='"&trim(arrBPL_Price(CNT1))&"', "
			SQL = SQL & "	Partner_P_Name='"&trim(arrPartner_P_Name(CNT1))&"', "
			SQL = SQL & "	BPL_Apply_Date='"&trim(arrBPL_Apply_Date(CNT1))&"' "
			SQL = SQL & "where BPL_Code='"&arrID_All(CNT1)&"' "
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
<form name="frmRedirect" action="bpl_list.asp" method=post>

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
<form name="frmRedirect" action="bpl_list.asp" method=post>

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