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
dim	arrP_P_No
dim	arrP_Work_Type
dim	arrP_Desc
dim	arrP_Qty
dim	arrP_Safe_Qty
dim	arrP_LGE_Price
dim	arrPartner_P_Price_1
dim	arrPartner_P_Name_1
dim	arrPartner_P_Price_2
dim	arrPartner_P_Name_2

arrID_All				= split(Request("strID_All")&" "			,", ")
arrP_P_No				= split(Request("P_P_No")&" "				,", ")
arrP_Work_Type			= split(Request("P_Work_Type")&" "			,", ")
arrP_Desc				= split(Request("P_Desc")&" "				,", ")
arrP_Qty				= split(Request("P_Qty")&" "				,", ")
arrP_Safe_Qty			= split(Request("P_Safe_Qty")&" "			,", ")
arrP_LGE_Price			= split(Request("P_LGE_Price")&" "			,", ")
arrPartner_P_Price_1	= split(Request("Partner_P_Price_1")&" "	,", ")
arrPartner_P_Name_1		= split(Request("Partner_P_Name_1")&" "		,", ")
arrPartner_P_Price_2	= split(Request("Partner_P_Price_2")&" "	,", ")
arrPartner_P_Name_2		= split(Request("Partner_P_Name_2")&" "		,", ")


set RS1 = Server.CreateObject("ADODB.RecordSet")
for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)				= trim(arrID_All(CNT1))
	arrP_P_No(CNT1)				= trim(arrP_P_No(CNT1))
	arrP_Work_Type(CNT1)		= trim(arrP_Work_Type(CNT1))
	arrP_Desc(CNT1)				= trim(arrP_Desc(CNT1))
	arrP_Qty(CNT1)				= trim(arrP_Qty(CNT1))
	arrP_Safe_Qty(CNT1)			= trim(arrP_Safe_Qty(CNT1))
	arrP_LGE_Price(CNT1)		= trim(arrP_LGE_Price(CNT1))
	arrPartner_P_Price_1(CNT1)	= trim(arrPartner_P_Price_1(CNT1))
	arrPartner_P_Name_1(CNT1)	= trim(arrPartner_P_Name_1(CNT1))
	arrPartner_P_Price_2(CNT1)	= trim(arrPartner_P_Price_2(CNT1))
	arrPartner_P_Name_2(CNT1)	= trim(arrPartner_P_Name_2(CNT1))
next

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트

	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		if strError_Temp = "" then
			SQL = "select top 1 P_Code from tbParts where P_P_No='"&arrP_P_No(CNT1)&"' and P_Code <> '"&arrID_All(CNT1)&"'"
			RS1.Open SQL,sys_DBCon
			if not(RS1.Eof or RS1.Bof) then
				strError_Temp = strError_Temp & "* "&arrID_All(CNT1)&"번 항목과 동일한 파트넘버의 아이템이 이미 등록되어있습니다.\n"
			end if
			RS1.Close
		end if		
		
		if strError_Temp = "" then
			SQL = 		"update tbParts set "
			SQL = SQL & "	P_P_No='"&arrP_P_No(CNT1)&"', "
			SQL = SQL & "	P_Work_Type='"&arrP_Work_Type(CNT1)&"', "
			SQL = SQL & "	P_Desc='"&arrP_Desc(CNT1)&"', "
			if isNumeric(arrP_Qty(CNT1)) then
				SQL = SQL & "	P_Qty='"&arrP_Qty(CNT1)&"', "
			end if
			if isNumeric(arrP_Safe_Qty(CNT1)) then
				SQL = SQL & "	P_Safe_Qty='"&arrP_Safe_Qty(CNT1)&"', "
			end if
			if isNumeric(arrP_LGE_Price(CNT1)) then
				SQL = SQL & "	P_LGE_Price='"&arrP_LGE_Price(CNT1)&"' "
			end if
			SQL = SQL & "where P_Code='"&arrID_All(CNT1)&"' "
			sys_DBCon.execute(SQL)
			
			if isNumeric(arrPartner_P_Price_1(CNT1)) then
				if arrPartner_P_Price_1(CNT1) > 0 and arrPartner_P_Name_1(CNT1) <> "" then
					SQL = "select PP_Code from tbParts_Price where Parts_P_P_No='"&arrP_P_No(CNT1)&"' and PP_Last_YN='Y'"
					RS1.Open SQL,sys_DBCon
					if RS1.Eof or RS1.Bof then	'최종사용단가가 없다. 최종사용단가 등록
						SQL = "insert into tbParts_Price (Parts_P_P_No, PP_Price, Partner_P_Name, PP_Last_YN) values ('"&arrP_P_No(CNT1)&"',"&arrPartner_P_Price_1(CNT1)&",'"&arrPartner_P_Name_1(CNT1)&"','Y')"
						sys_DBCon.execute(SQL)
					else						'최종사용단가가 있다. > 업체와 가격을 업데이트
						SQL = "update tbParts_Price set PP_Price = "&arrPartner_P_Price_1(CNT1)&", Partner_P_Name='"&arrPartner_P_Name_1(CNT1)&"' where PP_Code='"&RS1("PP_Code")&"'"
						sys_DBCon.execute(SQL)
					end if
					RS1.Close
				else
					SQL = "delete tbParts_Price where Parts_P_P_No='"&arrP_P_No(CNT1)&"' and PP_Last_YN='Y'"
					sys_DBCon.execute(SQL)
				end if
			end if
			
			if isNumeric(arrPartner_P_Price_2(CNT1)) then
				if arrPartner_P_Price_2(CNT1) > 0 and arrPartner_P_Name_2(CNT1) <> "" then
					SQL = "select PP_Code from tbParts_Price where Parts_P_P_No='"&arrP_P_No(CNT1)&"' and PP_Last_YN<>'Y'"
					RS1.Open SQL,sys_DBCon
					if RS1.Eof or RS1.Bof then	'보조단가가 없다. 보조단가 등록
						SQL = "insert into tbParts_Price (Parts_P_P_No, PP_Price, Partner_P_Name, PP_Last_YN) values ('"&arrP_P_No(CNT1)&"',"&arrPartner_P_Price_2(CNT1)&",'"&arrPartner_P_Name_2(CNT1)&"','')"
						sys_DBCon.execute(SQL)
					else						'보조단가가 있다. > 업체와 가격을 업데이트
						SQL = "update tbParts_Price set PP_Price = "&arrPartner_P_Price_2(CNT1)&", Partner_P_Name='"&arrPartner_P_Name_2(CNT1)&"' where PP_Code='"&RS1("PP_Code")&"'"
						sys_DBCon.execute(SQL)
					end if
					RS1.Close
				else
					SQL = "delete tbParts_Price where Parts_P_P_No='"&arrP_P_No(CNT1)&"' and PP_Last_YN=''"
					sys_DBCon.execute(SQL)
				end if
			end if
			
			'call Update_History("tbCritical_History","tbParts",Parts_MSE_PM_Code,"PM_Qty",arrPM_Qty(CNT1))
			'call Update_History("tbCritical_History","tbParts",Parts_MSE_PM_Code,"PM_MSE_Price_1",arrPM_MSE_Price_1(CNT1))
			'call Update_History("tbCritical_History","tbParts",Parts_MSE_PM_Code,"PM_MSE_Price_2",arrPM_MSE_Price_2(CNT1))
			'call Update_History("tbCommon_History","tbParts",arrID_All(CNT1),"PL_LGE_Price",arrPL_LGE_Price(CNT1))
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
<form name="frmRedirect" action="p_data_list.asp" method=post>

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
<form name="frmRedirect" action="p_data_list.asp" method=post>

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