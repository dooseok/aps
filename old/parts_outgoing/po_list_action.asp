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
dim arrPO_Out_Date
dim arrPO_Issued_Date
dim arrPO_Price
dim arrPO_Qty
dim arrPartner_P_Name
dim arrPO_State
dim arrPO_Payment_Type
dim arrPO_Remark

arrID_All				= split(Request("strID_All")&" "		,", ")
arrParts_P_P_No			= split(Request("Parts_P_P_No")&" "		,", ")
arrPO_Out_Date			= split(Request("PO_Out_Date")&" "		,", ")
arrPO_Price				= split(Request("PO_Price")&" "			,", ")
arrPO_Qty				= split(Request("PO_Qty")&" "			,", ")
arrPartner_P_Name		= split(Request("Partner_P_Name")&" "	,", ")
arrPO_State				= split(Request("PO_State")&" "			,", ")
arrPO_Payment_Type		= split(Request("PO_Payment_Type")&" "	,", ")
arrPO_Remark			= split(Request("PO_Remark")&" "		,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrParts_P_P_No(CNT1)		= trim(arrParts_P_P_No(CNT1))
	arrPO_Out_Date(CNT1)		= trim(arrPO_Out_Date(CNT1))
	arrPO_Price(CNT1)			= trim(arrPO_Price(CNT1))
	arrPO_Qty(CNT1)				= trim(arrPO_Qty(CNT1))
	arrPartner_P_Name(CNT1)		= trim(arrPartner_P_Name(CNT1))
	arrPO_State(CNT1)			= trim(arrPO_State(CNT1))
	arrPO_Payment_Type(CNT1)	= trim(arrPO_Payment_Type(CNT1))
	arrPO_Remark(CNT1)			= trim(arrPO_Remark(CNT1))
next


rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	

	for CNT1 = 0 to ubound(arrID_All)
		if arrPO_State(CNT1) = "출고준비" then
			strError_Temp = ""
			
			if strError_Temp = "" then
			end if
		
			if strError_Temp = "" then
				
				SQL = "select * from tbParts_Outgoing where PO_Code = '"&arrID_All(CNT1)&"'"
				RS1.Open SQL,sys_DBconString,3,2,&H0001
				with RS1
					.Fields("Parts_P_P_No")		= arrParts_P_P_No(CNT1)
					
					if arrPO_Out_Date(CNT1) <> "" then
						.Fields("PO_Out_Date")		= arrPO_Out_Date(CNT1)
					end if
					if arrPO_Price(CNT1) <> "" then
						.Fields("PO_Price")			= arrPO_Price(CNT1)
					end if
					if arrPO_Qty(CNT1) <> "" then
						.Fields("PO_Qty")			= arrPO_Qty(CNT1)
					end if
					if arrPartner_P_Name(CNT1) <> "" then
						.Fields("Partner_P_Name")	= arrPartner_P_Name(CNT1)
					end if
					if arrPO_State(CNT1) <> "" then
						.Fields("PO_State")			= arrPO_State(CNT1)
					end if
					if arrPO_Payment_Type(CNT1) <> "" then
						.Fields("PO_Payment_Type")	= arrPO_Payment_Type(CNT1)
					end if
					if arrPO_Remark(CNT1) <> "" then
						.Fields("PO_Remark")		= arrPO_Remark(CNT1)
					end if
					
					.Update
					.Close
				end with
				
				SQL = "select Partner_P_Name from tbParts_Outgoing where PO_Code = '"&arrID_All(CNT1)&"'"
				RS1.Open SQL,sys_DBCon
				if RS1.Eof or RS1.Bof then
				else
					SQL = "update tbPartner set P_Payment_Type = '"&arrPO_Payment_Type(CNT1)&"' where P_Name='"&RS1("Partner_P_Name")&"'"
					sys_DBCon.execute(SQL)
				end if
				RS1.Close
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
<form name="frmRedirect" action="po_list.asp" method=post>
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
<form name="frmRedirect" action="po_list.asp" method=post>
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