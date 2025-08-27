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
dim arrMP_Price
dim arrPartner_P_Name

dim MOD_In_Qty

arrID_All			= split(Request("strID_All")&" "		,", ")
arrMP_Price			= split(Request("MP_Price")&" "			,", ")
arrPartner_P_Name	= split(Request("Partner_P_Name")&" "	,", ")


set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrMP_Price(CNT1)			= trim(arrMP_Price(CNT1))
	arrPartner_P_Name(CNT1)		= trim(arrPartner_P_Name(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""

		if strError_Temp = "" then
			SQL = 		"update tbMaterial_Price set "
			SQL = SQL & "Material_M_P_No='"&Request("s_Material_M_P_No")&"', "
			SQL = SQL & "Partner_P_Name='"&arrPartner_P_Name(CNT1)&"', "
			SQL = SQL & "MP_Price="&arrMP_Price(CNT1)&" where MP_Code='"&arrID_All(CNT1)&"'"
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
<form name="frmRedirect" action="MP_list.asp" method=post>
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
<form name="frmRedirect" action="MP_list.asp" method=post>
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