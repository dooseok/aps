<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
rem ��������
dim RS1
dim SQL
dim CNT1

dim strError
dim strError_Temp

dim arrID_All
dim arrMO_Qty
dim arrMO_Due_Date
dim arrMO_Check_1_YN
dim arrMO_Check_2_YN
dim arrMO_Check_3_YN

dim strMO_Check_1_YN

arrID_All			= split(Request("strID_All")&" "		,", ")
arrMO_Qty			= split(Request("MO_Qty")&" "			,", ")
arrMO_Due_Date		= split(Request("MO_Due_Date")&" "		,", ")
arrMO_Check_1_YN	= split(Request("MO_Check_1_YN")&" "	,", ")
arrMO_Check_2_YN	= split(Request("MO_Check_2_YN")&" "	,", ")
arrMO_Check_3_YN	= split(Request("MO_Check_3_YN")&" "	,", ")

strMO_Check_1_YN	= Request("MO_Check_1_YN")
strMO_Check_1_YN	= replace(strMO_Check_1_YN,",","")
strMO_Check_1_YN	= replace(strMO_Check_1_YN," ","")
strMO_Check_1_YN	= ","&strMO_Check_1_YN&","

for CNT1 = 0 to ubound(arrID_All)
	arrID_All(CNT1)			= trim(arrID_All(CNT1))
	arrMO_Qty(CNT1)			= trim(arrMO_Qty(CNT1))
	arrMO_Due_Date(CNT1)	= trim(arrMO_Due_Date(CNT1))
next
response.write Request("MO_Check_1_YN")

set RS1 = Server.CreateObject("ADODB.RecordSet")

rem �����޼����� ���� ��� ����ȵ�
if strError = "" then	
	
	rem DB ������Ʈ
	for CNT1 = 0 to ubound(arrID_All)
		strError_Temp = ""
	
		SQL = "select MO_Check_1_YN,Material_M_P_No from tbMaterial_Order where MO_Code='"&arrID_All(CNT1)&"' "
		RS1.Open SQL,sys_DBCon
		if RS1("MO_Check_1_YN") = "" or isnull(RS1("MO_Check_1_YN")) then '���ֿϷ� üũ�� �ȵǾ� �־��� ���¶��
			if instr(strMO_Check_1_YN,","&arrID_All(CNT1)&",") > 0 then '�̹��� üũ�� �Ǿ��ٸ�, ������������ > �̷���� �ݿ�
				SQL = "update tbMaterial_Order set "
				SQL = SQL & "	MO_Qty="&arrMO_Qty(CNT1)&", " 
				SQL = SQL & "	MO_Due_Date='"&arrMO_Due_Date(CNT1)&"', "
				SQL = SQL & "	MO_Edit_Date='"&date()&"', "
				SQL = SQL & "	MO_Edit_ID='"&gM_ID&"' "
				SQL = SQL & "where MO_Code='"&arrID_All(CNT1)&"' "
				sys_DBCon.execute(SQL)
			
			else 'üũ�� �ȵǾ��ִٸ�, ������������ > �̷���� �ݿ�����
				SQL = "update tbMaterial_Order set "
				SQL = SQL & "	MO_Qty="&arrMO_Qty(CNT1)&", "
				SQL = SQL & "	MO_Due_Date='"&arrMO_Due_Date(CNT1)&"', "
				SQL = SQL & "	MO_Edit_Date='"&date()&"', "
				SQL = SQL & "	MO_Edit_ID='"&gM_ID&"' "
				SQL = SQL & "where MO_Code='"&arrID_All(CNT1)&"' "
				sys_DBCon.execute(SQL)
			
			end if
		else '�̹� ���ֿϷᰡ üũ�Ǿ��־��ٸ�, ���������Ұ� > �̷���� �ݿ� �Ұ�.
			SQL = "update tbMaterial_Order set "
			'SQL = SQL & "	MO_Qty="&arrMO_Qty(CNT1)&", "
			SQL = SQL & "	MO_Due_Date='"&arrMO_Due_Date(CNT1)&"', "
			SQL = SQL & "	MO_Edit_Date='"&date()&"', "
			SQL = SQL & "	MO_Edit_ID='"&gM_ID&"' "
			SQL = SQL & "where MO_Code='"&arrID_All(CNT1)&"' "
			sys_DBCon.execute(SQL)
			
		end if
		
		SQL = "update tbMaterial set M_Qty_Include_coming = (select sum(isnull(MO_Qty,0)-isnull(MO_Qty_In,0)) from tbMaterial_Order where Material_M_P_No = M_P_No) where M_P_No = '"&RS1("Material_M_P_No")&"'"
		sys_DBCon.execute(SQL) 
		RS1.Close
		
		
		strError = strError & strError_Temp
	next

	for CNT1 = 0 to ubound(arrMO_Check_1_YN)
		if isnumeric(arrMO_Check_1_YN(CNT1)) then
			SQL = "update tbMaterial_Order set MO_Check_1_YN = '"&gM_ID&"' where MO_Code = "& arrMO_Check_1_YN(CNT1)
			sys_DBCon.execute(SQL)
		end if
	next
	if gM_ID = "leeth" then
		for CNT1 = 0 to ubound(arrMO_Check_2_YN)
			if isnumeric(arrMO_Check_2_YN(CNT1)) then
				SQL = "update tbMaterial_Order set MO_Check_2_YN = '"&gM_ID&"' where MO_Code = "& arrMO_Check_2_YN(CNT1)
				sys_DBCon.execute(SQL)
			end if
		next
	end if
	if gM_ID = "shindk" then
		for CNT1 = 0 to ubound(arrMO_Check_3_YN)
			if isnumeric(arrMO_Check_3_YN(CNT1)) then
				SQL = "update tbMaterial_Order set MO_Check_3_YN = '"&gM_ID&"' where MO_Code = "& arrMO_Check_3_YN(CNT1)
				sys_DBCon.execute(SQL)
			end if
		next
	end if
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
<form name="frmRedirect" action="mo_list.asp" method=post>

<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
	'strError = strError & "* �Ϻ��� ������ ��ҵǾ����ϴ�."
%>
<form name="frmRedirect" action="mo_list.asp" method=post>

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