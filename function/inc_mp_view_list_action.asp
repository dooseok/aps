<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem 변수선언
dim RS1
dim SQL
dim CNT1
dim CNT2

dim strError
dim strError_Temp

dim MP_Process

dim arrLP_Code
dim arrMP_Work_Order
dim arrMP_Model
dim arrMP_D_No_1
dim arrMP_D_No_2
dim arrMP_D_No_3
dim arrMP_D_No_4

dim arrMP_Plan_Date
dim arrMP_Line
dim arrMP_Plan_Qty

MP_Process			= Request("MP_Process")
arrLP_Code			= split(Request("strLP_Code"),", ")
arrMP_Work_Order	= split(Request("strMP_Work_Order"),", ")
arrMP_Model			= split(Request("strMP_Model"),", ")
arrMP_D_No_1		= split(Request("arrMP_D_No_1"),", ")
arrMP_D_No_2		= split(Request("arrMP_D_No_2"),", ")
arrMP_D_No_3		= split(Request("arrMP_D_No_3"),", ")
arrMP_D_No_4		= split(Request("arrMP_D_No_4"),", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트

	for CNT1 = 0 to ubound(arrLP_Code)
		SQL = "delete tbMSE_Plan where MP_Process='"&MP_Process&"' and MP_Work_Order='"&arrMP_Work_Order&"'"
		sys_DBCon.execute(SQL)
			
		arrMP_Plan_Date	= split(Request(arrLP_Code(CNT1)&"_strMP_Plan_Date"),", ")
		arrMP_Line		= split(Request(arrLP_Code(CNT1)&"_strMP_Line"),", ")
		arrMP_Plan_Qty	= split(Request(arrLP_Code(CNT1)&"_strMP_Plan_Qty"),", ")
		
		for CNT2=0 to ubound(arrMP_Plan_Date)
		strError_Temp = ""
		
		if strError_Temp = "" then
			if arrMP_Plan_Date(CNT2) <> "" and isNumeric(arrMP_Plan_Qty(CNT2)) then
				SQL = "Insert into tbMSE_Plan (MP_Process,MP_Line,MP_Work_Order,MP_Model,MP_D_No_1,MP_D_No_2,MP_D_No_3,MP_D_No_4,MP_Plan_Date,MP_Plan_Qty) values ("
				SQL = SQL & "'"&MP_Process&"',"
				SQL = SQL & "'"&arrMP_Line(CNT2)&"',"
				SQL = SQL & "'"&arrMP_Work_Order(CNT1)&"',"
				SQL = SQL & "'"&arrMP_Model(CNT1)&"',"
				SQL = SQL & "'"&arrMP_D_No_1(CNT1)&"',"
				SQL = SQL & "'"&arrMP_D_No_2(CNT1)&"',"
				SQL = SQL & "'"&arrMP_D_No_3(CNT1)&"',"
				SQL = SQL & "'"&arrMP_D_No_4(CNT1)&"',"
				SQL = SQL & "'"&arrMP_Plan_Date(CNT2)&"',"
				SQL = SQL & arrMP_Plan_Qty(CNT2)&")"
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
<form name="frmRedirect" action="inc_mp_view_list.asp" method=post>
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
<form name="frmRedirect" action="inc_mp_view_list.asp" method=post>
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