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
dim arrB_Standard_Time

arrID_All			= split(Request("strID_All")&" ",", ")
arrB_Standard_Time	= split(Request("B_Standard_Time")&" ",", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			 = trim(arrID_All(CNT1))
	arrB_Standard_Time(CNT1) = trim(arrB_Standard_Time(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
	
		if strError_Temp = "" then
			SQL = 		"update tbBOM set "
			SQL = SQL & "B_Standard_Time='"&arrB_Standard_Time(CNT1)&"' where B_Code='"&arrID_All(CNT1)&"'"
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
<form name="frmRedirect" action="process_state_bom_st_list.asp" method=post>

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
<form name="frmRedirect" action="process_state_bom_st_list.asp" method=post>

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