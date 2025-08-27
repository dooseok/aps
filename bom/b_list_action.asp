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
dim arrB_Issue_Date
dim arrB_Tool
dim arrB_Desc
dim arrB_Spec

arrID_All		= split(Request("strID_All")&" "	,", ")
arrB_Issue_Date	= split(Request("B_Issue_Date")&" "	,", ")
arrB_Tool	= split(Request("B_Tool")&" "	,", ")
arrB_Desc	= split(Request("B_Desc")&" "	,", ")
arrB_Spec	= split(Request("B_Spec")&" "	,", ")
for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			= trim(arrID_All(CNT1))
	arrB_Issue_Date(CNT1)	= trim(arrB_Issue_Date(CNT1))
	arrB_Tool(CNT1)			= trim(arrB_Tool(CNT1))
	arrB_Desc(CNT1)			= trim(arrB_Desc(CNT1))
	arrB_Spec(CNT1)			= trim(arrB_Spec(CNT1))
next

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
				
		arrID_All(CNT1)				= trim(arrID_All(CNT1))
		arrB_Issue_Date(CNT1)		= trim(arrB_Issue_Date(CNT1))
		if strError_Temp = "" then
			SQL = "update tbBOM set "
			SQL = SQL & "	b_Issue_Date='"&arrB_Issue_Date(CNT1)&"', "
			SQL = SQL & "	b_Class='', "
			SQL = SQL & "	b_Tool='"&arrB_Tool(CNT1)&"', "
			SQL = SQL & "	b_Desc='"&arrB_Desc(CNT1)&"', "
			SQL = SQL & "	b_Spec='"&arrB_Spec(CNT1)&"' "
			SQL = SQL & "where B_Code='"&arrID_All(CNT1)&"' "
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
<form name="frmRedirect" action="b_list.asp" method=post>

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
<form name="frmRedirect" action="b_list.asp" method=post>

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