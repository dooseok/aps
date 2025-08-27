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

dim M_P_No

dim arrID_All
dim arrPartner_P_Name
dim arrM_P_No_Sub
dim arrM_OSP_YN
dim arrM_Desc
dim arrM_Spec
dim arrM_Price
dim arrM_Price_LGE
dim arrM_Price_Temp_YN
dim arrM_Price_Apply_Date
dim arrM_Process
dim arrM_Maker
dim arrM_Division

dim Partner_P_Name
dim old_MPL_Price
dim old_MPL_Price_LGE
dim old_MPL_Apply_Date
dim old_MPL_Temp_YN
		 
dim bChange_YN

arrID_All			= split(Request("strID_All")&" "		,", ")
arrM_P_No_Sub			= split(Request("M_P_No_Sub")&" "		,", ")
arrM_OSP_YN			= split(Request("M_OSP_YN")&" "			,", ")
arrM_Desc			= split(Request("M_Desc")&" "			,", ")
arrM_Spec			= split(Request("M_Spec")&" "			,", ")
arrM_Price			= split(Request("M_Price")&" "			,", ")
arrM_Price_LGE		= split(Request("M_Price_LGE")&" "		,", ")
arrM_Price_Temp_YN		= split(Request("M_Price_Temp_YN")&" "		,", ")
arrM_Price_Apply_Date		= split(Request("M_Price_Apply_Date")&" "		,", ")
arrM_Process		= split(Request("M_Process")&" "		,", ")
arrM_Maker			= split(Request("M_Maker")&" "			,", ")
arrM_Division		= split(Request("M_Division")&" "		,", ")

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			= trim(arrID_All(CNT1))
	arrM_P_No_Sub(CNT1)		= trim(arrM_P_No_Sub(CNT1))
	arrM_OSP_YN(CNT1)		= trim(arrM_OSP_YN(CNT1))
	arrM_Desc(CNT1)			= trim(arrM_Desc(CNT1))
	arrM_Spec(CNT1)			= trim(arrM_Spec(CNT1))
	arrM_Price(CNT1)		= trim(arrM_Price(CNT1))
	arrM_Price_LGE(CNT1)	= trim(arrM_Price_LGE(CNT1))
	arrM_Price_Temp_YN(CNT1)	= trim(arrM_Price_Temp_YN(CNT1))
	arrM_Price_Apply_Date(CNT1)	= trim(arrM_Price_Apply_Date(CNT1))
	arrM_Process(CNT1)		= trim(arrM_Process(CNT1))
	arrM_Maker(CNT1)		= trim(arrM_Maker(CNT1))
	arrM_Division(CNT1)		= trim(arrM_Division(CNT1))
next

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	
	rem DB 업데이트
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
		
		'변경된 항목이 있는 경우만 업데이트 하도록. (해당행의 폼값을 가져와서, DB를 조회해서, 동일한 내용이 나온다면, 방금 form submit으로 변경된 내용이 없는 것과 같다.)
		SQL = "select top 1 * from tbMaterial where "
		SQL = SQL & "	M_P_No_Sub='"&arrM_P_No_Sub(CNT1)&"' and "
		SQL = SQL & "	M_OSP_YN='"&arrM_OSP_YN(CNT1)&"' and "
		SQL = SQL & "	M_Desc='"&arrM_Desc(CNT1)&"' and "
		SQL = SQL & "	M_Spec='"&arrM_Spec(CNT1)&"' and "
		if instr(admin_material_handler,"-"&gM_ID&"-") > 0 then
			SQL = SQL & "	M_Price='"&cdbl(arrM_Price(CNT1))&"' and "
			SQL = SQL & "	M_Price_LGE='"&cdbl(arrM_Price_LGE(CNT1))&"' and "
			SQL = SQL & "	M_Price_Temp_YN='"&arrM_Price_Temp_YN(CNT1)&"' and "
			SQL = SQL & "	M_Price_Apply_Date='"&arrM_Price_Apply_Date(CNT1)&"' and "
		end if
		SQL = SQL & "	M_Process='"&arrM_Process(CNT1)&"' and "
		SQL = SQL & "	M_Maker='"&arrM_Maker(CNT1)&"' and "
		SQL = SQL & "	M_Division='"&arrM_Division(CNT1)&"' and "
		SQL = SQL & " M_Code='"&arrID_All(CNT1)&"' "
		RS1.Open SQL,sys_DBCon
		if not(RS1.Eof or RS1.Bof) then
			strError_Temp = "변동사항 없음"
		end if
		RS1.Close
		
		'업데이트 할 레코드의 파트넘버와 거래처를 가져온다.,
		SQL = "select M_P_No, Partner_P_Name from tbMaterial where M_Code = "&arrID_All(CNT1)
		RS1.Open SQL,sys_DBCon
		M_P_No = RS1("M_P_No")
		Partner_P_Name = RS1("Partner_P_Name")
		RS1.Close
		
		'해당 자재레코드 업데이트(딱 해당 레코드만 수정일과 단가 업데이트를 한다.)
		if strError_Temp = "" then
			SQL = "update tbMaterial set "
			SQL = SQL & "	M_P_No_Sub='"&arrM_P_No_Sub(CNT1)&"', "
			SQL = SQL & "	M_OSP_YN='"&arrM_OSP_YN(CNT1)&"', "
			SQL = SQL & "	M_Desc='"&arrM_Desc(CNT1)&"', "
			SQL = SQL & "	M_Spec='"&arrM_Spec(CNT1)&"', "
			if instr(admin_material_handler,"-"&gM_ID&"-") > 0 then
				SQL = SQL & "	M_Price='"&cdbl(arrM_Price(CNT1))&"', "
				SQL = SQL & "	M_Price_LGE='"&cdbl(arrM_Price_LGE(CNT1))&"', "
				SQL = SQL & "	M_Price_Temp_YN='"&arrM_Price_Temp_YN(CNT1)&"', "
				SQL = SQL & "	M_Price_Apply_Date='"&arrM_Price_Apply_Date(CNT1)&"', "
			end if
			SQL = SQL & "	M_Process='"&arrM_Process(CNT1)&"', "
			SQL = SQL & "	M_Maker='"&arrM_Maker(CNT1)&"', "
			SQL = SQL & "	M_Division='"&arrM_Division(CNT1)&"', "
			SQL = SQL & "	M_Edit_Date='"&date()&"', "
			SQL = SQL & "	M_Edit_ID='"&gM_ID&"' "
			SQL = SQL & "where M_Code='"&arrID_All(CNT1)&"' "
			sys_DBCon.execute(SQL)
		end if
		
		if instr(admin_material_handler,"-"&gM_ID&"-") > 0 then	
			'가져온 파트넘버-거래처 정보를 기준으로 직구입가 또는 인증가가 변한게 있다면,
			SQL = "select top 1 MPL_Price, MPL_Price_LGE, MPL_Apply_Date, MPL_Temp_YN  from tbMaterial_Price_Log where Material_M_P_No = '"&M_P_No&"' and Partner_P_Name = '"&Partner_P_Name&"' order by MPL_Code desc"
			RS1.Open SQL,sys_DBCon
			old_MPL_Price			= RS1("MPL_Price")
			old_MPL_Price_Lge	= RS1("MPL_Price_LGE")
			old_MPL_Apply_Date	= RS1("MPL_Apply_Date")
			old_MPL_Temp_YN	= RS1("MPL_Temp_YN")
			RS1.Close
			bChange_YN = "N"
			if cdbl(old_MPL_Price) <> cdbl(arrM_Price(CNT1)) then
				bChange_YN = "Y"
			end if
			if cdbl(old_MPL_Price_Lge) <> cdbl(arrM_Price_LGE(CNT1)) then
				bChange_YN = "Y"
			end if
			if cstr(old_MPL_Apply_Date) <> cstr(arrM_Price_Apply_Date(CNT1)) then
				bChange_YN = "Y"
			end if
			if cstr(old_MPL_Temp_YN) <> cstr(arrM_Price_Temp_YN(CNT1)) then
				bChange_YN = "Y"
			end if
			
			if bChange_YN = "Y" then
				SQL = "update tbMaterial_Order set MO_Price_Temp_YN = '"&arrM_Price_Temp_YN(CNT1)&"', MO_Price = "&cdbl(arrM_Price(CNT1))&" where (MO_Qty_In_Date >= '"&arrM_Price_Apply_Date(CNT1)&"') and Material_M_P_No = '"&M_P_No&"' and Partner_P_Name = '"&Partner_P_Name&"'"
				sys_DBCon.execute(SQL)
				SQL = "update tbMaterial_Order set MO_Price_Temp_YN = '"&arrM_Price_Temp_YN(CNT1)&"', MO_Price = "&cdbl(arrM_Price(CNT1))&" where (MO_Qty_In_Date is null and MO_Order_Date >= '"&arrM_Price_Apply_Date(CNT1)&"') and Material_M_P_No = '"&M_P_No&"' and Partner_P_Name = '"&Partner_P_Name&"'"
				sys_DBCon.execute(SQL)
				SQL = "update tbMaterial_Transaction set MT_Price = "&cdbl(arrM_Price(CNT1))&" where MT_Date >= '"&arrM_Price_Apply_Date(CNT1)&"' and Material_M_P_No = '"&M_P_No&"' and Partner_P_Name = '"&Partner_P_Name&"'"
				sys_DBCon.execute(SQL)
				
				SQL = "insert into tbMaterial_Price_Log (Material_M_P_No, Partner_P_Name, MPL_Price, MPL_Price_Diff, MPL_Price_Old, MPL_Price_LGE, MPL_Price_LGE_Diff, MPL_Price_LGE_Old, MPL_Apply_Date, MPL_Temp_YN, MPL_Reg_Date, MPL_Reg_ID, MPL_Desc) values "
				SQL = SQL & "('"&M_P_No&"','"&Partner_P_Name&"',"&cdbl(arrM_Price(CNT1))&","&cdbl(arrM_Price(CNT1))-cdbl(old_MPL_Price)&","&cdbl(old_MPL_Price)&","&cdbl(arrM_Price_LGE(CNT1))&","&cdbl(arrM_Price_LGE(CNT1))-cdbl(old_MPL_Price_LGE)&","&cdbl(old_MPL_Price_LGE)&",'"&arrM_Price_Apply_Date(CNT1)&"','"&arrM_Price_Temp_YN(CNT1)&"','"&date()&"','"&gM_ID&"','')"
				sys_DBCon.execute(SQL)
			end if
		end if
		
		strError = strError & strError_Temp
	next
end if
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
<form name="frmRedirect" action="m_list.asp" method=post>

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
<form name="frmRedirect" action="m_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
//alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->