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
dim arrMT_Out_byWho
dim arrMT_Desc

arrID_All			= split(Request("strID_All")&" "	,", ")
arrMT_Out_byWho		= split(Request("MT_Out_byWho")&" "		,", ")
arrMT_Desc			= split(Request("MT_Desc")&" "		,", ")


for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrMT_Out_byWho(CNT1)		= trim(arrMT_Out_byWho(CNT1))
	arrMT_Desc(CNT1)			= trim(arrMT_Desc(CNT1))
next

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		if strError_Temp = "" then
			SQL = "update tbMaterial_Transaction set "
			SQL = SQL & "	MT_Out_byWho='"&arrMT_Out_byWho(CNT1)&"', "
			SQL = SQL & "	MT_Desc='"&arrMT_Desc(CNT1)&"' "
			SQL = SQL & "where MT_Code='"&arrID_All(CNT1)&"' "
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
<form name="frmRedirect" action="mt_list_out.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
	strError = strError & "* 일부의 수정이 취소되었습니다."
%>
<form name="frmRedirect" action="mt_list_out.asp" method=post>

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