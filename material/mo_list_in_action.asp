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
dim arrMO_Qty_In
dim arrMO_Qty_In_Date
dim arrMO_Qty_In_Desc
dim MO_Qty_In_ID

dim MO_Qty_In_Date

dim Material_M_P_No
dim Partner_P_Name
dim MO_Qty_In
dim MO_Price

dim M_Qty

arrID_All			= split(Request("strID_All")&" "		,", ")
arrMO_Qty_In		= split(Request("MO_Qty_In")&" "		,", ")
arrMO_Qty_In_Date	= split(Request("MO_Qty_In_Date")&" "	,", ")
arrMO_Qty_In_Desc	= split(Request("MO_Qty_In_Desc")&" "	,", ")
MO_Qty_In_ID		= gM_ID

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			= trim(arrID_All(CNT1))
	arrMO_Qty_In(CNT1)		= trim(arrMO_Qty_In(CNT1))
next
set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
	
		if strError_Temp = "" and arrMO_Qty_In(CNT1) <> 0 then '수량이 변동된 내용만 내용 업데이트

			SQL = "select MO_Qty_In, Material_M_P_No, Partner_P_Name, MO_Price, MO_Qty_In_Date from tbMaterial_Order where MO_Code='"&arrID_All(CNT1)&"' "
			RS1.Open SQL,sys_DBCon
			Material_M_P_No	= RS1("Material_M_P_No")
			Partner_P_Name	= RS1("Partner_P_Name")
			MO_Qty_In		= RS1("MO_Qty_In")
			MO_Price		= RS1("MO_Price")
			MO_Qty_In_Date = RS1("MO_Qty_In_Date")
			RS1.Close

			SQL = "select M_Qty from tbMaterial where M_P_No='"&Material_M_P_No&"' "
			RS1.Open SQL,sys_DBCon
			M_Qty = RS1("M_Qty")
			RS1.Close

			if isnumeric(MO_Qty_In) then
			else
					MO_Qty_In = 0
			end if
			if clng(MO_Qty_In) = 0 or MO_Qty_In_Date = "1900-01-01" then '기존에 입고량이 없었다면 이것저것 수정가능

				SQL = "update tbMaterial_Order set "
				SQL = SQL & "	MO_Qty_In="&arrMO_Qty_In(CNT1)&", "
				if arrMO_Qty_In_Date(CNT1) = "" or isnull(arrMO_Qty_In_Date(CNT1)) then
					arrMO_Qty_In_Date(CNT1) = date()
				end if
				SQL = SQL & "	MO_Qty_In_Date='"&arrMO_Qty_In_Date(CNT1)&"', "
				SQL = SQL & "	MO_Qty_In_ID='"&MO_Qty_In_ID&"', "
				SQL = SQL & "	MO_Edit_Date='"&date()&"', "
				SQL = SQL & "	MO_Edit_ID='"&gM_ID&"' "
				SQL = SQL & "where MO_Code='"&arrID_All(CNT1)&"'"
				sys_DBCon.execute(SQL)
				
				'입고량 만큼 현재재고 증가
				SQL = "update tbMaterial set M_Qty = M_Qty + "&arrMO_Qty_In(CNT1)&" where M_P_No = '"&Material_M_P_No&"'"
				sys_DBCon.execute(SQL)
				
				RS1.Open "tbMaterial_Transaction",sys_DBConString,3,2,2
				with RS1
					.AddNew
					.Fields("Material_M_P_No")		= Material_M_P_No
					.Fields("Partner_P_Name")		= Partner_P_Name
					.Fields("MT_Out_byWho")			= ""
					.Fields("MT_Date")				= cstr(arrMO_Qty_In_Date(CNT1))
					.Fields("MT_Price")				= MO_Price
					.Fields("MT_Qty_In")			= arrMO_Qty_In(CNT1)
					.Fields("MT_Qty_Out")			= 0
					.Fields("MT_Qty_Update")		= 0
					.Fields("MT_Qty_Last")			= M_Qty
					.Fields("MT_Qty_Now")			= M_Qty + arrMO_Qty_In(CNT1)
					.Fields("MT_Desc")				= "입고"
					.Fields("MT_Reg_Date")			= now()
					.Fields("MT_Reg_ID")			= gM_ID
					.Update	
					.Close
				end with
			else '입고량에 관계없이 비고 항목은 업데이트
				SQL = "update tbMaterial_Order set "
				SQL = SQL & "	MO_Qty_In_Desc='"&arrMO_Qty_In_Desc(CNT1)&"' "
				SQL = SQL & "where MO_Code='"&arrID_All(CNT1)&"'"
				sys_DBCon.execute(SQL) 
			end if
			
			SQL = "update tbMaterial set M_Qty_Include_coming = (select sum(isnull(MO_Qty,0)-isnull(MO_Qty_In,0)) from tbMaterial_Order where Material_M_P_No = M_P_No) where M_P_No = '"&Material_M_P_No&"'"
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
<form name="frmRedirect" action="mo_list_in.asp" method=post>

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
<form name="frmRedirect" action="mo_list_in.asp" method=post>

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