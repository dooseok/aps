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
dim arrParts_P_P_No
dim arrParts_M_Process
dim arrPOD_Qty
dim arrPOD_in_Qty
dim arrPOD_In_Date
dim arrPOD_Due_Date
dim arrPOD_Price
dim arrPOD_Remark
dim Partner_P_Name

dim POD_In_Qty

dim M_Qty

arrID_All				= split(Request("strID_All")&" "			,", ")
arrParts_P_P_No		= split(Request("Parts_P_P_No")&" "		,", ")
arrParts_M_Process	= split(Request("Parts_M_Process")&" "	,", ")
arrPOD_Qty				= split(Request("POD_Qty")&" "				,", ")
arrPOD_In_Qty			= split(Request("POD_In_Qty")&" "			,", ")
arrPOD_In_Date			= split(Request("POD_In_Date")&" "			,", ")
arrPOD_Due_Date			= split(Request("POD_Due_Date")&" "			,", ")
arrPOD_Price			= split(Request("POD_Price")&" "			,", ")
arrPOD_Remark			= split(Request("POD_Remark")&" "			,", ")
Partner_P_Name			= Request("s_Partner_P_Name")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrParts_P_P_No(CNT1)	= trim(arrParts_P_P_No(CNT1))
	arrParts_M_Process(CNT1)	= trim(arrParts_M_Process(CNT1))
	arrPOD_Qty(CNT1)			= trim(arrPOD_Qty(CNT1))
	arrPOD_In_Qty(CNT1)			= trim(arrPOD_In_Qty(CNT1))
	arrPOD_In_Date(CNT1)		= trim(arrPOD_In_Date(CNT1))
	arrPOD_Due_Date(CNT1)		= trim(arrPOD_Due_Date(CNT1))
	arrPOD_Price(CNT1)			= trim(arrPOD_Price(CNT1))
	arrPOD_Remark(CNT1)			= trim(arrPOD_Remark(CNT1))
	
	if arrPOD_Qty(CNT1) = "" then
		arrPOD_Qty(CNT1) = 0
	end if
	if arrPOD_Price(CNT1) = "" then
		arrPOD_Price(CNT1) = 0
	end if
	if arrPOD_In_Qty(CNT1) = "" then
		arrPOD_In_Qty(CNT1) = 0
	end if
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	for CNT1 = 0 to ubound(arrID_All)
		if arrParts_P_P_No(CNT1) <> "" then
			strError_Temp = ""
	
			if strError_Temp = "" then
				set RS1 = Server.CreateObject("ADODB.RecordSet")
				SQL = "select POD_In_Qty from tbParts_Order_Detail where POD_Code='"&arrID_All(CNT1)&"'"
				RS1.Open SQL,sys_DBCon
				POD_In_Qty = 0
				if not(RS1.Eof or RS1.Bof) then
					if isnumeric(RS1("POD_In_Qty")) then
						POD_In_Qty = RS1("POD_In_Qty")
					end if
				end if
				RS1.Close
	
	
				SQL = 		"update tbParts_Order_Detail set "
				SQL = SQL & "Parts_P_P_No='"&arrParts_P_P_No(CNT1)&"', "
				SQL = SQL & "POD_Qty='"&arrPOD_Qty(CNT1)&"', "
				SQL = SQL & "POD_In_Qty='"&arrPOD_In_Qty(CNT1)&"', "
				if arrPOD_In_Date(CNT1) <> "" then
					SQL = SQL & "POD_In_Date='"&arrPOD_In_Date(CNT1)&"', "
				end if
				if arrPOD_Due_Date(CNT1) <> "" then
					SQL = SQL & "POD_Due_Date='"&arrPOD_Due_Date(CNT1)&"', "
				end if
				SQL = SQL & "POD_Remark='"&arrPOD_Remark(CNT1)&"', "
				SQL = SQL & "POD_Price='"&arrPOD_Price(CNT1)&"' where POD_Code='"&arrID_All(CNT1)&"'"
				sys_DBCon.execute(SQL)
				
				SQL = 		"update tbParts set "
				SQL = SQL & "M_Qty=M_Qty - "&POD_In_Qty&" + "&arrPOD_in_Qty(CNT1)&", "
				SQL = SQL & "M_Process='"&arrParts_M_Process(CNT1)&"' where P_P_No='"&arrParts_P_P_No(CNT1)&"'"
	
				SQL = "select * from tbParts where P_P_No='"&arrParts_P_P_No(CNT1)&"'"
				RS1.Open SQL,sys_DBCon
				M_Qty = RS1("M_Qty")
				RS1.Close
				
				SQL = "insert into tbParts_Stock_History (Parts_P_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
				SQL = SQL &arrParts_P_P_No(CNT1)&"',"
				SQL = SQL &arrPOD_in_Qty(CNT1) - POD_In_Qty&","
				SQL = SQL &M_Qty&",'"
				SQL = SQL &"발주입고','"
				SQL = SQL &date()&"','"
				SQL = SQL &Request("s_Partner_P_Name")&"')"
				if arrPOD_in_Qty(CNT1) - POD_In_Qty <> 0 then
					sys_DBCon.execute(SQL)
				end if
	
				'동일한 거래처, 파트넘버의 가격데이타 조회
				SQL = "select MP_Price from tbParts_Price where Parts_P_P_No ='"&arrParts_P_P_No(CNT1)&"' and Partner_P_Name='"&Request("s_Partner_P_Name")&"'"
				RS1.Open SQL,sys_DBCon
				if RS1.Eof or RS1.Bof then '없으면 추가
					SQL = 		"insert into tbParts_Price (Parts_P_P_No, Partner_P_Name, MP_Price) values "
					SQL = SQL & "('"&arrParts_P_P_No(CNT1)&"','"&Request("s_Partner_P_Name")&"',"&arrPOD_Price(CNT1)&")"
					sys_DBCon.execute(SQL)
				else '있으면 없데이트
					SQL = 		"update tbParts_Price set "
					SQL = SQL & "MP_Price = "&arrPOD_Price(CNT1)&" where Parts_P_P_No ='"&arrParts_P_P_No(CNT1)&"' and Partner_P_Name='"&Request("s_Partner_P_Name")&"'"
					sys_DBCon.execute(SQL)
				end if
				RS1.Close
	
				set RS1 = nothing
			end if
	
			strError = strError & strError_Temp
		end if
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
<form name="frmRedirect" action="POD_list.asp" method=post>
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
<form name="frmRedirect" action="POD_list.asp" method=post>
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