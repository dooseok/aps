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
dim arrLM_Company
dim arrLM_Name
dim arrBOM_Sub_BS_D_No_1
dim arrBOM_Sub_BS_D_No_2
dim arrBOM_Sub_BS_D_No_3
dim arrBOM_Sub_BS_D_No_4

arrID_All				= split(Request("strID_All")&" "			,", ")
arrLM_Company			= split(Request("LM_Company")&" "			,", ")
arrLM_Name				= split(Request("LM_Name")&" "				,", ")
arrBOM_Sub_BS_D_No_1	= split(Request("BOM_Sub_BS_D_No_1")&" "	,", ")
arrBOM_Sub_BS_D_No_2	= split(Request("BOM_Sub_BS_D_No_2")&" "	,", ")
arrBOM_Sub_BS_D_No_3 	= split(Request("BOM_Sub_BS_D_No_3")&" "	,", ")
arrBOM_Sub_BS_D_No_4	= split(Request("BOM_Sub_BS_D_No_4")&" "	,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)					= trim(arrID_All(CNT1))
	arrLM_Company(CNT1)				= trim(arrLM_Company(CNT1))
	arrLM_Name(CNT1)				= trim(arrLM_Name(CNT1))
	arrBOM_Sub_BS_D_No_1(CNT1)		= trim(arrBOM_Sub_BS_D_No_1(CNT1))
	arrBOM_Sub_BS_D_No_2(CNT1)		= trim(arrBOM_Sub_BS_D_No_2(CNT1))
	arrBOM_Sub_BS_D_No_3(CNT1)		= trim(arrBOM_Sub_BS_D_No_3(CNT1))
	arrBOM_Sub_BS_D_No_4(CNT1)		= trim(arrBOM_Sub_BS_D_No_4(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		if strError_Temp = "" then
			SQL = "select top 1 LM_Name from tbLGE_Model where LM_Name='"&arrLM_Name(CNT1)&"' and LM_Code <> '"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if not(RS1.Eof or RS1.Bof) then
				strError_Temp = strError_Temp & "* "&arrID_All(CNT1)&"번 항목과 동일한 모델정보가 이미 등록되어있습니다.\n"
			end if
			RS1.Close
		end if
	
		if strError_Temp = "" then
			SQL = 		"update tbLGE_Model set "
			SQL = SQL & "LM_Company='"&arrLM_Company(CNT1)&"', "			
			SQL = SQL & "LM_Name='"&arrLM_Name(CNT1)&"', "			
				
			SQL = SQL & "BOM_Sub_BS_D_No_1='"&arrBOM_Sub_BS_D_No_1(CNT1)&"', "
			SQL = SQL & "BOM_Sub_BS_D_No_2='"&arrBOM_Sub_BS_D_No_2(CNT1)&"', "
			SQL = SQL & "BOM_Sub_BS_D_No_3='"&arrBOM_Sub_BS_D_No_3(CNT1)&"', "
			SQL = SQL & "BOM_Sub_BS_D_No_4='"&arrBOM_Sub_BS_D_No_4(CNT1)&"' where LM_Code='"&arrID_All(CNT1)&"'"
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
<form name="frmRedirect" action="lm_list.asp" method=post>
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
<form name="frmRedirect" action="lm_list.asp" method=post>
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