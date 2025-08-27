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
dim arrPI_To_Date
dim arrPI_In_Date
dim arrPI_Price
dim arrPI_Qty
dim arrPartner_P_Name
dim arrPI_State
dim arrPI_Payment_Type
dim arrPI_Remark

dim Exist_YN
dim Exist_Partner_YN

arrID_All				= split(Request("strID_All")&" "		,", ")
arrParts_P_P_No			= split(Request("Parts_P_P_No")&" "		,", ")
arrPI_To_Date			= split(Request("PI_To_Date")&" "		,", ")
arrPI_In_Date			= split(Request("PI_In_Date")&" "		,", ")
arrPI_Price				= split(Request("PI_Price")			,", ")
arrPI_Qty				= split(Request("PI_Qty")			,", ")
arrPartner_P_Name		= split(Request("Partner_P_Name")	,", ")
arrPI_State				= split(Request("PI_State")&" "			,", ")
arrPI_Payment_Type		= split(Request("PI_Payment_Type")&" "	,", ")
arrPI_Remark			= split(Request("PI_Remark")&" "		,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrParts_P_P_No(CNT1)		= trim(arrParts_P_P_No(CNT1))
	arrPI_To_Date(CNT1)			= trim(arrPI_To_Date(CNT1))
	arrPI_In_Date(CNT1)			= trim(arrPI_In_Date(CNT1))
	arrPI_Price(CNT1)			= trim(arrPI_Price(CNT1))
	arrPI_Qty(CNT1)				= trim(arrPI_Qty(CNT1))
	arrPartner_P_Name(CNT1)		= trim(arrPartner_P_Name(CNT1))
	arrPI_State(CNT1)			= trim(arrPI_State(CNT1))
	arrPI_Payment_Type(CNT1)	= trim(arrPI_Payment_Type(CNT1))
	arrPI_Remark(CNT1)			= trim(arrPI_Remark(CNT1))
next


rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	

	for CNT1 = 0 to ubound(arrID_All)
		if arrPI_State(CNT1) = "발주준비" then
			strError_Temp = ""
			
			if strError_Temp = "" then
			end if
		
			if strError_Temp = "" then
				
				SQL = "select * from tbParts_Incoming where PI_Code = '"&arrID_All(CNT1)&"'"
				RS1.Open SQL,sys_DBconString,3,2,&H0001
				with RS1
					.Fields("Parts_P_P_No")		= arrParts_P_P_No(CNT1)
					
					if arrPI_To_Date(CNT1) <> "" then
						.Fields("PI_To_Date")		= arrPI_To_Date(CNT1)
					end if
					if arrPI_In_Date(CNT1) <> "" then
						.Fields("PI_In_Date")		= arrPI_In_Date(CNT1)
					end if
					if arrPI_Price(CNT1) <> "" then
						.Fields("PI_Price")			= arrPI_Price(CNT1)
					end if
					if arrPI_Qty(CNT1) <> "" then
						.Fields("PI_Qty")			= arrPI_Qty(CNT1)
					end if
					if arrPartner_P_Name(CNT1) <> "" then
						.Fields("Partner_P_Name")	= arrPartner_P_Name(CNT1)
					end if
					if arrPI_State(CNT1) <> "" then
						.Fields("PI_State")			= arrPI_State(CNT1)
					end if
					if arrPI_Payment_Type(CNT1) <> "" then
						.Fields("PI_Payment_Type")	= arrPI_Payment_Type(CNT1)
					end if
					if arrPI_Remark(CNT1) <> "" then
						.Fields("PI_Remark")		= arrPI_Remark(CNT1)
					end if
					
					.Update
					.Close
				end with
				
				if Exist_YN = "Y" and Exist_Partner_YN = "Y" then		'단가정보 있고, 해당거래처도 있다. -> 단가만 업데이트
					'기존 데이터를 전부 N으로 바꾼다.
					SQL = "update tbParts_Price set PP_Last_YN = '' where Parts_P_P_No = '"&arrParts_P_P_No(CNT1)&"'"
					sys_DBCon.execute(SQL)
					'같은 거래처의 정보를 업데이트 한다.
					SQL = "update tbParts_Price set PP_Last_YN = 'Y', PP_Price = "&arrPI_Price(CNT1)&" where Parts_P_P_No = '"&arrParts_P_P_No(CNT1)&"' and Partner_P_Name='"&arrPartner_P_Name(CNT1)&"'"
					sys_DBCon.execute(SQL)
				elseif Exist_YN = "Y" and Exist_Partner_YN = "N" then	'단가정보 있고, 해당거래처가 없다.
					'기존 데이터를 전부 N으로 바꾼다.
					SQL = "update tbParts_Price set PP_Last_YN = '' where Parts_P_P_No = '"&arrParts_P_P_No(CNT1)&"'"
					sys_DBCon.execute(SQL)
					'새로운 데이터를 추가한다.
					SQL = "insert into tbParts_Price (Parts_P_P_No, Partner_P_Name, PP_Price, PP_Last_YN) values ('"&arrParts_P_P_No(CNT1)&"','"&arrPartner_P_Name(CNT1)&"',"&arrPI_Price(CNT1)&",'Y')"
					sys_DBCon.execute(SQL)
				elseif Exist_YN = "N" and Exist_Partner_YN = "N" then	'단가정보 없고, 해당거래처도 없다. -> 새로운 단가정보 추가
					'새로운 단가정보를 등록한다.
					SQL = "insert into tbParts_Price (Parts_P_P_No, Partner_P_Name, PP_Price, PP_Last_YN) values ('"&arrParts_P_P_No(CNT1)&"','"&arrPartner_P_Name(CNT1)&"',"&arrPI_Price(CNT1)&",'Y')"
					sys_DBCon.execute(SQL)
				end if
				
				SQL = "update tbPartner set P_Payment_Type = '"&arrPI_Payment_Type(CNT1)&"' where P_Name='"&arrPartner_P_Name(CNT1)&"'"
				sys_DBCon.execute(SQL)
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
<form name="frmRedirect" action="pi_list.asp" method=post>
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
<form name="frmRedirect" action="pi_list.asp" method=post>
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