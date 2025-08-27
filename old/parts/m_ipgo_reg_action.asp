<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1
dim CNT1

dim arrPartner_P_Name
dim arrMaterial_M_P_No
dim arrMOD_In_Date
dim arrMOD_In_Qty

dim MO_Code

dim M_Qty

dim temp
dim strError
dim strError_Temp

dim URL_Prev
dim URL_Next

arrPartner_P_Name		= split(Request("Partner_P_Name")&" "	,", ")
arrMaterial_M_P_No		= split(Request("Material_M_P_No")&" "	,", ")
arrMOD_In_Date			= split(Request("MOD_In_Date")&" "		,", ")
arrMOD_In_Qty			= split(Request("MOD_In_Qty")&" "		,", ")

for CNT1 = 0 to ubound(arrMaterial_M_P_No)
	arrPartner_P_Name(CNT1)		= trim(arrPartner_P_Name(CNT1))
	arrMaterial_M_P_No(CNT1)	= trim(arrMaterial_M_P_No(CNT1))
	arrMOD_In_Date(CNT1)		= trim(arrMOD_In_Date(CNT1))
	arrMOD_In_Qty(CNT1)			= trim(arrMOD_In_Qty(CNT1))
	
	if arrMOD_In_Date(CNT1) = "" then
		arrMOD_In_Date(CNT1) = date()
	end if
next

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem 에러메세지가 있을 경우 실행안됨
set RS1 = Server.CreateObject("ADODB.RecordSet")

if strError = "" then

	for CNT1 = ubound(arrMaterial_M_P_No) to 0
		strError_Temp = ""
		
		if arrMaterial_M_P_No(CNT1) <> "" then
			if strError_Temp = "" then
				SQL = "insert into tbMaterial_Order (MO_Date,MO_State,Partner_P_Name,MO_Due_Date) values "
				SQL = SQL & "('"&date()&"','입고완료','"&arrPartner_P_Name(CNT1)&"','"&arrMOD_In_Date(CNT1)&"')"
				sys_DBCon.execute(SQL)
			
				SQL = "select max(MO_Code) from tbMaterial_Order"
				RS1.Open SQL,sys_DBCon
				MO_Code = RS1(0)
				RS1.Close
				
			
				SQL = "insert into tbMaterial_Order_Detail (Material_Order_MO_Code,Material_M_P_No,MOD_Price,MOD_Qty,MOD_Remark,MOD_In_Qty,MOD_In_Date) values "
				SQL = SQL & "("&MO_Code&",'"&arrMaterial_M_P_No(CNT1)&"','',"&arrMOD_In_Qty(CNT1)&",'',"&arrMOD_In_Qty(CNT1)&",'"&arrMOD_In_Date(CNT1)&"')"
				sys_DBCon.execute(SQL)
				
				if arrMOD_In_Qty(CNT1) <> "0" then
					SQL = "select * from tbMaterial where M_P_No='"&arrMaterial_M_P_No(CNT1)&"'"
					RS1.Open SQL,sys_DBCon
					M_Qty = RS1("M_Qty")
					RS1.Close
					
					SQL = "insert into tbMaterial_Stock_History (Material_M_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
					SQL = SQL &arrMaterial_M_P_No(CNT1)&"',"
					SQL = SQL &arrMOD_In_Qty(CNT1)&","
					SQL = SQL &M_Qty + arrMOD_In_Qty(CNT1)&",'"
					SQL = SQL &"발주입고','"
					SQL = SQL &arrMOD_In_Date(CNT1)&"','"
					SQL = SQL &Request("s_Partner_P_Name")&"')"
					if arrMOD_In_Qty(CNT1) <> 0 then
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
	'strError = strError & "* 일부의 등록이 실패되었습니다."
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