<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim Material_M_P_No
dim Material_M_Desc
dim Material_M_Spec
dim MO_Qty

dim Partner_P_Name
dim MO_Price

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

Material_M_P_No	= trim(Request("Material_M_P_No"))
MO_Qty			= trim(Request("MO_Qty"))

set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트
	SQL = "select * from tbMaterial where M_P_No = '"&Material_M_P_No&"'"
	RS1.Open SQL,sys_DBCon
	Material_M_P_No	= RS1("M_P_No")
		Partner_P_Name	= RS1("Partner_P_Name")
		MO_Price		= RS1("M_Price")
		Material_M_Desc		= RS1("M_Desc")
		Material_M_Spec		= RS1("M_Spec")
	RS1.Close
	SQL = "insert tbMaterial_Order (Material_M_P_No,Material_M_Desc,Material_M_Spec,Partner_P_Name,MO_Price,MO_Qty,MO_Qty_In,O_Due_Date,MO_Order_Date,MO_Check_1_YN,MO_Check_2_YN,MO_Check_3_YN,MO_Reg_Date,MO_Reg_ID) values "
	SQL = SQL & "	('"&Material_M_P_No&"', "
	SQL = SQL & "	'"&Material_M_Desc&"', "
	SQL = SQL & "	'"&Material_M_Spec&"', "
	SQL = SQL & "	'"&Partner_P_Name&"', "
	SQL = SQL & "	'"&MO_Price&"', "
	if isnumeric(MO_Qty) then
	else
		MO_Qty = 0
	end if
	SQL = SQL & "	"&MO_Qty&", "
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

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="mo_list.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="mo_list.asp" method=post>

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