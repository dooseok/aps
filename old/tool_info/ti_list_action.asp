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
dim arrTI_Name
dim arrTI_Type

arrID_All				= split(Request("strID_All")&" "	,", ")
arrTI_Name				= split(Request("TI_Name")&" "		,", ")
arrTI_Type				= split(Request("TI_Type")&" "		,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)		= trim(arrID_All(CNT1))
	arrTI_Name(CNT1)	= trim(arrTI_Name(CNT1))
	arrTI_Type(CNT1)	= trim(arrTI_Type(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		if strError_Temp = "" then
			SQL = "select top 1 TI_Name from tbTool_Info where TI_Name='"&arrTI_Name(CNT1)&"' and TI_Code <> '"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if not(RS1.Eof or RS1.Bof) then
				strError_Temp = strError_Temp & "* "&arrID_All(CNT1)&"번 항목과 동일한 툴이 이미 등록되어있습니다.\n"
			end if
			RS1.Close
		end if
		
		
	
		if strError_Temp = "" then
			SQL = 		"update tbTool_Info set "
			SQL = SQL & "TI_Name='"&arrTI_Name(CNT1)&"', "			
			SQL = SQL & "TI_Type='"&arrTI_Type(CNT1)&"' where TI_Code='"&arrID_All(CNT1)&"'"
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
<form name="frmRedirect" action="ti_list.asp" method=post>
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
<form name="frmRedirect" action="ti_list.asp" method=post>
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