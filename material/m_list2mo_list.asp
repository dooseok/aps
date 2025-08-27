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



dim Material_M_P_No
dim Material_M_Desc
dim Material_M_Spec
dim Partner_P_Name
dim MO_Price
dim M_Price_Temp_YN

dim arrID

arrID		= split(Request("strID")&" " ,", ")

set RS1 = Server.CreateObject("ADODB.RecordSet")
if Request("strID") <> "" then
	for CNT1 = 0 to ubound(arrID)
		strError_Temp = ""
		SQL = "select * from tbMaterial where M_Code = " & arrID(CNT1)
		RS1.Open SQL,sys_DBcon
		Material_M_P_No	= RS1("M_P_No")
		Material_M_Desc	= replace(RS1("M_Desc"),"'","''")
		Material_M_Spec	= replace(RS1("M_Spec"),"'","''")
		Partner_P_Name	= RS1("Partner_P_Name")
		MO_Price		= RS1("M_Price")
		M_Price_Temp_YN		= RS1("M_Price_Temp_YN")
		RS1.Close
		
		if strError_Temp = "" then
			SQL = "insert tbMaterial_Order (Material_M_P_No,Material_M_Desc,Material_M_Spec,Partner_P_Name,MO_Price,MO_Price_Temp_YN,MO_Qty,MO_Qty_In,MO_Due_Date,MO_Order_Date,MO_Check_1_YN,MO_Check_2_YN,MO_Check_3_YN,MO_Reg_Date,MO_Reg_ID) values "
			SQL = SQL & "	('"&Material_M_P_No&"', "
			SQL = SQL & "	'"&Material_M_Desc&"', "
			SQL = SQL & "	'"&Material_M_Spec&"', "
			SQL = SQL & "	'"&Partner_P_Name&"', "
			SQL = SQL & "	'"&MO_Price&"', "
			SQL = SQL & "	'"&M_Price_Temp_YN&"', "
			SQL = SQL & "	0, "
			SQL = SQL & "	0, "
			SQL = SQL & "	'"&date()&"', "
			SQL = SQL & "	'"&date()&"', "
			SQL = SQL & "	'', "
			SQL = SQL & "	'', "
			SQL = SQL & "	'', "
			SQL = SQL & "	'"&date()&"', "
			SQL = SQL & "	'"&gM_ID&"') "
			sys_DBCon.execute(SQL)
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
parent.frmMoFrame.cols='100px,*';
parent.frameMain.location.reload();
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
alert("<%=strError%>");
frmRedirect.submit();
</script>
<%
end if
%>



<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->