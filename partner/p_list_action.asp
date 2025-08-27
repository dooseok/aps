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
dim arrP_Code
dim arrP_Name
dim arrP_Business_No
dim arrP_Owner
dim arrP_Pay_Method
dim arrP_Memo
dim arrP_Email


arrID_All			= split(Request("strID_All")&" "		,", ")
arrP_Name			= split(Request("P_Name")&" "			,", ")
arrP_Business_No	= split(Request("P_Business_No")&" "	,", ")
arrP_Owner			= split(Request("P_Owner")&" "			,", ")
arrP_Pay_Method		= split(Request("P_Pay_Method")&" "		,", ")
arrP_Memo			= split(Request("P_Memo")&" "			,", ")
arrP_Email			= split(Request("P_Email")&" "			,", ")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			= trim(arrID_All(CNT1))
	arrP_Name(CNT1)			= trim(arrP_Name(CNT1))
	arrP_Business_No(CNT1)	= trim(arrP_Business_No(CNT1))
	arrP_Owner(CNT1)		= trim(arrP_Owner(CNT1))
	arrP_Pay_Method(CNT1)	= trim(arrP_Pay_Method(CNT1))
	arrP_Memo(CNT1)			= trim(arrP_Memo(CNT1))
	arrP_Email(CNT1)		= trim(arrP_Email(CNT1))
next

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		if strError_Temp = "" then
			SQL = "select top 1 P_Code from tbPartner where P_Name='"&arrP_Name(CNT1)&"' and P_Code <> '"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if not(RS1.Eof or RS1.Bof) then
				strError_Temp = strError_Temp & "* "&arrID_All(CNT1)&"번 항목과 동일한 상호며의 거래처가 이미 등록되어있습니다.\n"
			end if
			RS1.Close
		end if
	
		if strError_Temp = "" then
			SQL = "update tbPartner set "
			SQL = SQL & "	P_Name='"&arrP_Name(CNT1)&"', "
			SQL = SQL & "	P_Business_No='"&arrP_Business_No(CNT1)&"', "
			SQL = SQL & "	P_Owner='"&arrP_Owner(CNT1)&"', "
			SQL = SQL & "	P_Pay_Method='"&arrP_Pay_Method(CNT1)&"', "
			SQL = SQL & "	P_Memo='"&arrP_Memo(CNT1)&"', "
			SQL = SQL & "	P_Email='"&arrP_Email(CNT1)&"' "
			SQL = SQL & "where P_Code='"&arrID_All(CNT1)&"' "
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
<form name="frmRedirect" action="p_list.asp" method=post>

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
<form name="frmRedirect" action="p_list.asp" method=post>

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