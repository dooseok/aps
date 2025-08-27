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
dim arrMaterial_M_P_No
dim arrMaterial_M_Process
dim arrMTD_Qty
dim arrMTD_Remark
dim arrMTD_Ipgo_Date

dim MTD_Qty
dim MT_Company

dim M_Qty

arrID_All				= split(Request("strID_All")&" "			,", ")
arrMaterial_M_P_No		= split(Request("Material_M_P_No")&" "		,", ")
arrMaterial_M_Process	= split(Request("Material_M_Process")&" "	,", ")
arrMTD_Qty				= split(Request("MTD_Qty")&" "				,", ")
arrMTD_Remark			= split(Request("MTD_Remark")&" "			,", ")
arrMTD_Ipgo_Date		= split(Request("MTD_Ipgo_Date")&" "		,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrMaterial_M_P_No(CNT1)	= trim(arrMaterial_M_P_No(CNT1))
	arrMaterial_M_Process(CNT1)	= trim(arrMaterial_M_Process(CNT1))
	arrMTD_Qty(CNT1)			= trim(arrMTD_Qty(CNT1))
	arrMTD_Remark(CNT1)			= trim(arrMTD_Remark(CNT1))
	arrMTD_Ipgo_Date(CNT1)		= trim(arrMTD_Ipgo_Date(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""

		if strError_Temp = "" then
			
			SQL = "select MTD_Qty from tbMaterial_Transaction_Detail where MTD_Code='"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			MTD_Qty = 0
			if not(RS1.Eof or RS1.Bof) then
				if isnumeric(RS1("MTD_Qty")) then
					MTD_Qty = RS1("MTD_Qty")
				end if
			end if
			RS1.Close
			

			SQL = 		"update tbMaterial_Transaction_Detail set "
			SQL = SQL & "Material_M_P_No='"&arrMaterial_M_P_No(CNT1)&"', "
			SQL = SQL & "MTD_Qty='"&arrMTD_Qty(CNT1)&"', "
			
			if arrMTD_Ipgo_Date(CNT1) <> "" then
				SQL = SQL & "MTD_Ipgo_Date='"&arrMTD_Ipgo_Date(CNT1)&"', "
			end if
			
			SQL = SQL & "MTD_Remark='"&arrMTD_Remark(CNT1)&"' where MTD_Code='"&arrID_All(CNT1)&"'"
			sys_DBCon.execute(SQL)

			if Request("s_IpgoOrChulgo") = "Ipgo" then
				SQL = "update tbMaterial set M_Qty=M_Qty - "&MTD_Qty&" + "&arrMTD_Qty(CNT1)&", M_Process='"&arrMaterial_M_Process(CNT1)&"' where M_P_No='"&arrMaterial_M_P_No(CNT1)&"'"
			elseif Request("s_IpgoOrChulgo") = "Chulgo" then
				SQL = "update tbMaterial set M_Qty=M_Qty + "&MTD_Qty&" - "&arrMTD_Qty(CNT1)&", M_Process='"&arrMaterial_M_Process(CNT1)&"' where M_P_No='"&arrMaterial_M_P_No(CNT1)&"'"
			end if
			sys_DBCon.execute(SQL)
			
			SQL = "select top 1 MT_Company from tbMaterial_Transaction where MT_Code = "&Request("s_Material_Transaction_MT_Code")
			RS1.Open SQL,sys_DBCon
			MT_Company = RS1("MT_Company")
			RS1.Close
			
			SQL = "select top 1 M_Qty from tbMaterial where M_P_No='"&arrMaterial_M_P_No(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			M_Qty = RS1("M_Qty")
			RS1.Close
			
			SQL = "insert into tbMaterial_Stock_History (Material_M_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
			SQL = SQL &arrMaterial_M_P_No(CNT1)&"',"
			if Request("s_IpgoOrChulgo") = "Ipgo" then
				SQL = SQL &arrMTD_Qty(CNT1)-MTD_Qty&","
				SQL = SQL &M_Qty&",'"
				SQL = SQL &"사내입고수정','"
			elseif Request("s_IpgoOrChulgo") = "Chulgo" then
				SQL = SQL &MTD_Qty-arrMTD_Qty(CNT1)&","
				SQL = SQL &M_Qty&",'"
				SQL = SQL &"사내출고수정','"
			end if
			SQL = SQL &date()&"','"
			SQL = SQL &MT_Company&"')"
			
			if arrMTD_Qty(CNT1)-MTD_Qty = 0 and Request("s_IpgoOrChulgo") = "Ipgo" then
			elseif MTD_Qty-arrMTD_Qty(CNT1) = 0 and Request("s_IpgoOrChulgo") = "Chulgo" then
			else
				sys_DBCon.execute(SQL)
			end if
		end if

		strError = strError & strError_Temp
	next
	set RS1 = nothing
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
<form name="frmRedirect" action="MTD_list.asp" method=post>
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
<form name="frmRedirect" action="MTD_list.asp" method=post>
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