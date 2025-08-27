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
dim arrMOD_In_Date
dim arrMOD_In_Qty

dim Material_M_P_No
dim MOD_In_Qty

arrID_All		= split(Request("strID_All")&" "	,", ")
arrMOD_In_Date	= split(Request("MOD_In_Date")&" "	,", ")
arrMOD_In_Qty	= split(Request("MOD_In_Qty")&" "	,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)		= trim(arrID_All(CNT1))
	arrMOD_In_Date(CNT1)= trim(arrMOD_In_Date(CNT1))
	arrMOD_In_Qty(CNT1)	= trim(arrMOD_In_Qty(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""

		if strError_Temp = "" then
			
			SQL = "select Material_M_P_No, MOD_In_Qty from tbMaterial_Order_Detail where MOD_Code='"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			Material_M_P_No	= RS1("Material_M_P_No")
			MOD_In_Qty			= RS1("MOD_In_Qty")
			RS1.Close	
			
			SQL = 		"update tbMaterial_Order_Detail set "
			if arrMOD_In_Date(CNT1) = "" then
				SQL = SQL & "MOD_In_Date='"&date()&"', "
			else
				SQL = SQL & "MOD_In_Date='"&arrMOD_In_Date(CNT1)&"', "
			end if
			SQL = SQL & "MOD_In_Qty="&arrMOD_In_Qty(CNT1)&" where MOD_Code='"&arrID_All(CNT1)&"'"
			sys_DBCon.execute(SQL)
		
			
			SQL = "select * from tbMaterial where M_P_No='"&Material_M_P_No&"'"
			RS1.Open SQL,sys_DBCon
			M_Qty = RS1("M_Qty")
			RS1.Close
			
			SQL = "insert into tbMaterial_Stock_History (Material_M_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
			SQL = SQL &Material_M_P_No&"',"
			SQL = SQL &arrMOD_in_Qty(CNT1) - MOD_In_Qty&","
			SQL = SQL &M_Qty&",'"
			SQL = SQL &"발주입고','"
			SQL = SQL &arrMOD_In_Date(CNT1)&"','"
			SQL = SQL &Request("s_Partner_P_Name")&"')"
			if arrMOD_in_Qty(CNT1) - MOD_In_Qty <> 0 then
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
<form name="frmRedirect" action="M_ipgo_list.asp" method=post>

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
<form name="frmRedirect" action="M_ipgo_list.asp" method=post>

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