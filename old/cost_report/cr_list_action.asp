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
dim arrCR_Part
dim arrCR_Line
dim arrCR_Title
dim arrCR_Use1
dim arrCR_Use2
dim arrCR_Amount
dim arrCR_Date
dim arrCR_Memo

dim oldM_Qty

arrID_All		= split(Request("strID_All")&" "	,", ")
arrCR_Part		= split(Request("CR_Part")&" "		,", ")
arrCR_Line		= split(Request("CR_Line")&" "		,", ")
arrCR_Title		= split(Request("CR_Title")&" "		,", ")
arrCR_Use1		= split(Request("CR_Use1")&" "		,", ")
arrCR_Use2		= split(Request("CR_Use2")&" "		,", ")
arrCR_Amount	= split(Request("CR_Amount")&" "	,", ")
arrCR_Date		= split(Request("CR_Date")&" "		,", ")
arrCR_Memo		= split(Request("CR_Memo")&" "		,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)		= trim(arrID_All(CNT1))
	arrCR_Part(CNT1)	= trim(arrCR_Part(CNT1))
	arrCR_Line(CNT1)	= trim(arrCR_Line(CNT1))
	arrCR_Title(CNT1)	= trim(arrCR_Title(CNT1))
	arrCR_Use1(CNT1)	= trim(arrCR_Use1(CNT1))
	arrCR_Use2(CNT1)	= trim(arrCR_Use2(CNT1))
	arrCR_Amount(CNT1)	= trim(arrCR_Amount(CNT1))
	arrCR_Date(CNT1)	= trim(arrCR_Date(CNT1))
	arrCR_Memo(CNT1)	= trim(arrCR_Memo(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""

		if strError_Temp = "" then			
			SQL = 		"update tbCost_Report set "
			SQL = SQL & "CR_Part='"&arrCR_Part(CNT1)&"', "
			SQL = SQL & "CR_Line='"&arrCR_Line(CNT1)&"', "
			SQL = SQL & "CR_Title='"&arrCR_Title(CNT1)&"', "
			SQL = SQL & "CR_Use1='"&arrCR_Use1(CNT1)&"', "
			SQL = SQL & "CR_Use2='"&arrCR_Use2(CNT1)&"', "
			SQL = SQL & "CR_Amount="&arrCR_Amount(CNT1)&", "
			SQL = SQL & "CR_Date='"&arrCR_Date(CNT1)&"', "
			SQL = SQL & "CR_Memo='"&arrCR_Memo(CNT1)&"' where CR_Code='"&arrID_All(CNT1)&"'"
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
<form name="frmRedirect" action="CR_list.asp" method=post>
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
<form name="frmRedirect" action="CR_list.asp" method=post>
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