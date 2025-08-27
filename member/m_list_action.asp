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
dim arrM_ID
dim arrM_Channel
dim arrM_Password
dim arrM_Part
dim arrM_Position
dim arrM_Name
dim arrM_Email_1
dim arrM_HP
dim arrM_Enter_Date
dim arrM_Retire_Date
dim arrM_Authority
dim arrM_Use_YN

arrID_All			= split(Request("strID_All")&" "	,", ")
arrM_ID				= split(Request("M_ID")&" "			,", ")
arrM_Channel		= split(Request("M_Channel")&" "	,", ")
arrM_Password		= split(Request("M_Password")&" "	,", ")
arrM_Part			= split(Request("M_Part")&" "		,", ")
arrM_Position		= split(Request("M_Position")&" "	,", ")
arrM_Name			= split(Request("M_Name")&" "		,", ")
arrM_Email_1		= split(Request("M_Email_1")&" "	,", ")
arrM_HP				= split(Request("M_HP")&" "			,", ")
arrM_Enter_Date		= split(Request("M_Enter_Date")&" "	,", ")
arrM_Retire_Date	= split(Request("M_Retire_Date")&" ",", ")
arrM_Authority		= split(Request("M_Authority")&" "	,", ")
arrM_Use_YN			= split(Request("M_Use_YN")&" "		,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			= trim(arrID_All(CNT1))
	arrM_ID(CNT1)			= trim(arrM_ID(CNT1))
	arrM_Channel(CNT1)		= trim(arrM_Channel(CNT1))
	arrM_Password(CNT1)		= trim(arrM_Password(CNT1))
	arrM_Part(CNT1)			= trim(arrM_Part(CNT1))
	arrM_Position(CNT1)		= trim(arrM_Position(CNT1))
	arrM_Name(CNT1)			= trim(arrM_Name(CNT1))
	arrM_Email_1(CNT1)		= trim(arrM_Email_1(CNT1))
	arrM_HP(CNT1)			= trim(arrM_HP(CNT1))
	arrM_Enter_Date(CNT1)	= trim(arrM_Enter_Date(CNT1))
	arrM_Retire_Date(CNT1)	= trim(arrM_Retire_Date(CNT1))
	arrM_Authority(CNT1)	= trim(arrM_Authority(CNT1))
	arrM_Use_YN(CNT1)		= trim(arrM_Use_YN(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		if strError_Temp = "" then
			SQL = "select top 1 M_Code from tbMember where M_ID='"&arrM_ID(CNT1)&"' and M_Code <> '"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if not(RS1.Eof or RS1.Bof) then
				strError_Temp = strError_Temp & "* "&arrID_All(CNT1)&"번 항목과 동일한 아이디의 사원이 이미 등록되어있습니다.\n"
			end if
			RS1.Close
		end if
	
		if strError_Temp = "" then
			SQL = 		"update tbMember set "
			SQL = SQL & "M_ID='"&arrM_ID(CNT1)&"', "
			SQL = SQL & "M_Channel='"&arrM_Channel(CNT1)&"', "
			SQL = SQL & "M_Password='"&arrM_Password(CNT1)&"', "
			SQL = SQL & "M_Part='"&arrM_Part(CNT1)&"', "
			SQL = SQL & "M_Position='"&arrM_Position(CNT1)&"', "
			SQL = SQL & "M_Name='"&arrM_Name(CNT1)&"', "
			SQL = SQL & "M_Email_1='"&arrM_Email_1(CNT1)&"', "
			SQL = SQL & "M_HP='"&arrM_HP(CNT1)&"', "
			SQL = SQL & "M_Use_YN='"&arrM_Use_YN(CNT1)&"', "
			SQL = SQL & "M_Enter_Date='"&arrM_Enter_Date(CNT1)&"', "
			
			if arrM_Retire_Date(CNT1) <> "" then
				SQL = SQL & "M_Retire_Date='"&arrM_Retire_Date(CNT1)&"', "
			end if
			
			SQL = SQL & "M_Authority='"&arrM_Authority(CNT1)&"' where M_Code='"&arrID_All(CNT1)&"'"
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
<form name="frmRedirect" action="m_list.asp" method=post>

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
<form name="frmRedirect" action="m_list.asp" method=post>

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