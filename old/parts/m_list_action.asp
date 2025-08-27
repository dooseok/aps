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
dim arrM_P_No
dim arrM_Spec
dim arrM_Desc
dim arrM_Additional_Info
dim arrM_Qty
dim arrM_Process

dim oldM_Qty

arrID_All				= split(Request("strID_All")&" "		,", ")
arrM_P_No				= split(Request("M_P_No")&" "			,", ")
arrM_Spec				= split(Request("M_Spec")&" "			,", ")
arrM_Desc				= split(Request("M_Desc")&" "			,", ")
arrM_Additional_Info	= split(Request("M_Additional_Info")&" ",", ")
arrM_Qty				= split(Request("M_Qty")&" "			,", ")
arrM_Process			= split(Request("M_Process")&" "		,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrM_P_No(CNT1)				= trim(arrM_P_No(CNT1))
	arrM_Spec(CNT1)				= trim(arrM_Spec(CNT1))
	arrM_Desc(CNT1)				= trim(arrM_Desc(CNT1))
	arrM_Additional_Info(CNT1)	= trim(arrM_Additional_Info(CNT1))
	arrM_Qty(CNT1)				= trim(arrM_Qty(CNT1))
	arrM_Process(CNT1)			= trim(arrM_Process(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""

		if strError_Temp = "" then
			SQL = "select top 1 M_Code,M_Qty from tbMaterial where M_P_No='"&arrM_P_No(CNT1)&"' and M_Code <> '"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if not(RS1.Eof or RS1.Bof) then
				strError_Temp = strError_Temp & "* "&arrID_All(CNT1)&"번 항목과 동일한 파트넘버의 자재가 이미 등록되어있습니다.\n"
			end if
			RS1.Close
		end if
		

		if strError_Temp = "" then
			SQL = "select M_Qty from tbMaterial where M_Code='"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			oldM_Qty = RS1("M_Qty")
			RS1.Close
			
			SQL = 		"update tbMaterial set "
			SQL = SQL & "M_Spec='"&arrM_Spec(CNT1)&"', "
			SQL = SQL & "M_Desc='"&arrM_Desc(CNT1)&"', "
			SQL = SQL & "M_P_No='"&arrM_P_No(CNT1)&"', "
			SQL = SQL & "M_Additional_Info='"&arrM_Additional_Info(CNT1)&"', "
			SQL = SQL & "M_Qty="&arrM_Qty(CNT1)&", "
			SQL = SQL & "M_Process='"&arrM_Process(CNT1)&"' where M_Code='"&arrID_All(CNT1)&"'"
			sys_DBCon.execute(SQL)
			
			SQL = "insert into tbMaterial_Stock_History (Material_M_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
			SQL = SQL &arrM_P_No(CNT1)&"',"
			SQL = SQL &arrM_Qty(CNT1)-oldM_Qty&","
			SQL = SQL &arrM_Qty(CNT1)&",'"
			SQL = SQL &"직접수정','"
			SQL = SQL &date()&"','"
			SQL = SQL &"엠에스이')"
			
			if arrM_Qty(CNT1)-oldM_Qty <> 0 then
				sys_DBCon.execute(SQL)
			end if
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
<form name="frmRedirect" action="M_list.asp" method=post>

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
<form name="frmRedirect" action="M_list.asp" method=post>

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