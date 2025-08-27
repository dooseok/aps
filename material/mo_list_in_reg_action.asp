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
dim MO_Qty_In
dim MO_Qty_In_Date
dim MO_Qty_In_Desc

dim Partner_P_Name
dim MO_Price
dim M_Qty

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

Material_M_P_No	= trim(Request("Material_M_P_No"))
MO_Qty			= trim(Request("MO_Qty"))
MO_Qty_In		= trim(Request("MO_Qty_In"))
MO_Qty_In_Date	= Request("MO_Qty_In_Date")
MO_Qty_In_Desc	= trim(Request("MO_Qty_In_Desc"))


set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then	
	rem DB 업데이트
	SQL = "select Partner_P_Name,M_Price,M_Qty, M_Desc, M_Spec from tbMaterial where M_P_No = '"&Material_M_P_No&"'"
	RS1.Open SQL,sys_DBCon
	Partner_P_Name	= RS1("Partner_P_Name")	
	MO_Price		= RS1("M_Price")
	M_Qty			= RS1("M_Qty")
	Material_M_Desc			= RS1("M_Desc")
	Material_M_Spec			= RS1("M_Spec")
	RS1.Close
	
	SQL = "insert tbMaterial_Order (Material_M_P_No,Material_M_Desc,Material_M_Spec,Partner_P_Name,MO_Price,MO_Qty,MO_Qty_In,MO_Qty_In_ID,MO_Qty_In_Date,MO_Qty_In_Desc,MO_Due_Date,MO_Order_Date,MO_Check_1_YN,MO_Check_2_YN,MO_Check_3_YN,MO_Reg_Date,MO_Reg_ID) values "
	SQL = SQL & "	('"&Material_M_P_No&"', "
	SQL = SQL & "	('"&Material_M_Desc&"', "
	SQL = SQL & "	('"&Material_M_Spec&"', "
	SQL = SQL & "	'"&Partner_P_Name&"', "
	SQL = SQL & "	'"&MO_Price&"', "
	if isnumeric(MO_Qty) then
	else
		MO_Qty = 0
	end if
	SQL = SQL & "	"&MO_Qty&", "
	if isnumeric(MO_Qty_In) then
	else
		MO_Qty_In = 0
	end if
	SQL = SQL & "	"&MO_Qty_In&", "
	SQL = SQL & "	'"&gM_ID&"', "
	if MO_Qty_In_Date = "" or isnull(MO_Qty_In_Date) then
		MO_Qty_In_Date = date()
	end if
	SQL = SQL & "	'"&MO_Qty_In_Date&"', "
	SQL = SQL & "	'"&MO_Qty_In_Desc&"', "
	SQL = SQL & "	'"&date()&"', "
	SQL = SQL & "	'"&date()&"', "
	SQL = SQL & "	'"&gM_ID&"', "
	SQL = SQL & "	'', "
	SQL = SQL & "	'', "
	SQL = SQL & "	'"&date()&"', "
	SQL = SQL & "	'"&gM_ID&"') "
	sys_DBCon.execute(SQL)

	if MO_Qty_In = 0 then '발주량 > 0	입고량 = 0 then 
		
	else '발주량 > 0	입고량 > 0 then 

		'입고량 만큼 현재재고 증가
		SQL = "update tbMaterial set M_Qty = M_Qty + "&MO_Qty_In&" where M_P_No = '"&Material_M_P_No&"'"
		sys_DBCon.execute(SQL)
		
		RS1.Open "tbMaterial_Transaction",sys_DBConString,3,2,2
		with RS1
			.AddNew
			.Fields("Material_M_P_No")		= Material_M_P_No
			.Fields("Partner_P_Name")		= Partner_P_Name
			.Fields("MT_Out_byWho")			= ""
			.Fields("MT_Date")				= MO_Qty_In_Date
			.Fields("MT_Price")				= MO_Price
			.Fields("MT_Qty_In")			= MO_Qty_In
			.Fields("MT_Qty_Out")			= 0
			.Fields("MT_Qty_Update")		= 0
			.Fields("MT_Qty_Last")			= M_Qty
			.Fields("MT_Qty_Now")			= M_Qty + MO_Qty_In
			.Fields("MT_Desc")				= "입고"
			.Fields("MT_Reg_Date")			= now()
			.Fields("MT_Reg_ID")			= gM_ID
			.Update	
			.Close
		end with
	end if

	SQL = "update tbMaterial set M_Qty_Include_coming = (select sum(isnull(MO_Qty,0)-isnull(MO_Qty_In,0)) from tbMaterial_Order where Material_M_P_No = M_P_No) where M_P_No = '"&Material_M_P_No&"'"
	sys_DBCon.execute(SQL) 
end if

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="mo_list_in.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="mo_list_in.asp" method=post>

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