<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1
dim CNT1

dim arrParts_P_P_No
dim arrPOD_Qty
dim arrPOD_Price
dim arrPOD_Remark
dim arrPOD_Due_Date
dim arrPOD_In_Date
dim arrPOD_In_Qty

dim temp
dim strError
dim strError_Temp
dim URL_Prev
dim URL_Next

dim M_Qty

arrParts_P_P_No= split(Request("Parts_P_P_No")&" ",", ")
arrPOD_Qty				= split(Request("POD_Qty")&" "				,", ")
arrPOD_Price			= split(Request("POD_Price")&" "			,", ")
arrPOD_Remark			= split(Request("POD_Remark")&" "			,", ")
arrPOD_Due_Date		= split(Request("POD_Due_Date")&" "		,", ")
arrPOD_In_Date		= split(Request("POD_In_Date")&" "		,", ")
arrPOD_In_Qty			= split(Request("POD_In_Qty")&" "			,", ")

for CNT1 = 0 to ubound(arrParts_P_P_No)
	arrParts_P_P_No(CNT1)	= trim(arrParts_P_P_No(CNT1))
	arrPOD_Qty(CNT1)					= trim(arrPOD_Qty(CNT1))
	arrPOD_Price(CNT1)				= trim(arrPOD_Price(CNT1))
	arrPOD_Remark(CNT1)				= trim(arrPOD_Remark(CNT1))
	arrPOD_Due_Date(CNT1)			= trim(arrPOD_Due_Date(CNT1))
	arrPOD_In_Date(CNT1)			= trim(arrPOD_In_Date(CNT1))
	arrPOD_In_Qty(CNT1)				= trim(arrPOD_In_Qty(CNT1))
	
	if arrPOD_Qty(CNT1) = "" then
		arrPOD_Qty(CNT1) = 0
	end if
	
	if arrPOD_Due_Date(CNT1) = "" then
		arrPOD_Due_Date(CNT1) = date()
	end if
	
	if arrPOD_In_Qty(CNT1) = "" then
		arrPOD_In_Qty(CNT1) = 0
	end if
next

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

set RS1 = server.CreateObject("ADODB.RecordSet")
if strError = "" then

	for CNT1 = 0 to ubound(arrParts_P_P_No)
		strError_Temp = ""
		
		if arrParts_P_P_No(CNT1) <> "" then
			if strError_Temp = "" then
				'동일한 거래처, 파트넘버의 가격데이타 조회
				SQL = "update tbParts_Price set MP_Price = "&arrPOD_Price(CNT1)&" where Parts_P_P_No ='"&arrParts_P_P_No(CNT1)&"' and Partner_P_Name='"&Request("s_Partner_P_Name")&"'"
				sys_DBCon.execute(SQL)
			
				if arrPOD_In_Date(CNT1) = "" then
					SQL = "insert into tbParts_Order_Detail (Parts_Order_PO_Code,Parts_P_P_No,POD_Price,POD_Qty,POD_Remark,POD_Due_Date,POD_In_Qty) values "
					SQL = SQL & "("&Request("s_Parts_Order_PO_Code")&",'"&arrParts_P_P_No(CNT1)&"',"&arrPOD_Price(CNT1)&","&arrPOD_Qty(CNT1)&",'"&arrPOD_Remark(CNT1)&"','"&arrPOD_Due_Date(CNT1)&"',"&arrPOD_In_Qty(CNT1)&")"
				else
					SQL = "insert into tbParts_Order_Detail (Parts_Order_PO_Code,Parts_P_P_No,POD_Price,POD_Qty,POD_Remark,POD_Due_Date,POD_In_Date,POD_In_Qty) values "
					SQL = SQL & "("&Request("s_Parts_Order_PO_Code")&",'"&arrParts_P_P_No(CNT1)&"',"&arrPOD_Price(CNT1)&","&arrPOD_Qty(CNT1)&",'"&arrPOD_Remark(CNT1)&"','"&arrPOD_Due_Date(CNT1)&"','"&arrPOD_In_Date(CNT1)&"',"&arrPOD_In_Qty(CNT1)&")"
				end if
				sys_DBCon.execute(SQL)
			
				if arrPOD_In_Qty(CNT1) <> "0" then
							
					SQL = "select * from tbParts where P_P_No='"&arrParts_P_P_No(CNT1)&"'"
					RS1.Open SQL,sys_DBCon
					M_Qty = RS1("M_Qty")
					RS1.Close
				
					SQL = "insert into tbParts_Stock_History (Parts_P_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
					SQL = SQL &arrParts_P_P_No(CNT1)&"',"
					SQL = SQL &arrPOD_in_Qty(CNT1)&","
					SQL = SQL &M_Qty + arrPOD_in_Qty(CNT1)&",'"
					SQL = SQL &"발주입고','"
					SQL = SQL &arrPOD_in_Date(CNT1)&"','"
					SQL = SQL &Request("s_Partner_P_Name")&"')"
					if arrPOD_in_Qty(CNT1) <> 0 then
						sys_DBCon.execute(SQL)
					end if
					
				end if
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
	'strError = strError & "* 일부의 등록이 실패되었습니다."
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