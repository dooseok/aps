<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim M_P_No
dim M_P_No_Sub
dim Partner_P_Name
dim M_OSP_YN
dim M_Desc
dim M_Spec
dim M_Price_Temp_YN
dim M_Price
dim M_Price_LGE
dim M_Process
dim M_Maker
dim M_Division
dim M_Package_Unit

dim M_Qty
dim M_Qty_Safe
dim M_Qty_Include_coming


dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")

M_P_No			= trim(Request("M_P_No"))
M_P_No_Sub		= trim(Request("M_P_No_Sub"))
Partner_P_Name	= trim(Request("Partner_P_Name"))
M_OSP_YN		= trim(Request("M_OSP_YN"))
M_Desc			= trim(Request("M_Desc"))
M_Spec			= trim(Request("M_Spec"))
M_Price_Temp_YN			= trim(Request("M_Price_Temp_YN"))
M_Price			= trim(Request("M_Price"))
M_Price_LGE		= trim(Request("M_Price_LGE"))
M_Process		= trim(Request("M_Process"))
M_Maker			= trim(Request("M_Maker"))
M_Division		= trim(Request("M_Division"))

set RS1 = Server.CreateObject("ADODB.RecordSet")
rem 에러메세지가 있을 경우 실행안됨

'동일한 파트넘버와 거래처가 있다면 모두 실행 안함.
SQL = "select top 1 * from tbMaterial where M_P_No = '"&M_P_No&"' and Partner_P_Name = '"&Partner_P_Name&"'"
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
else
	strError = "동일한 파트넘버-거래처 조합이 이미 등록되어있습니다."
end if
RS1.Close


if strError = "" then	
	SQL = "select * from tbMaterial where M_P_No = '"&M_P_No&"'"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		M_Qty 								= 0
		M_Qty_Safe						= 0
		M_Qty_Include_coming	= 0
		M_Package_Unit				= 1
	else
		M_Qty									= RS1("M_Qty")
		M_Qty_Safe						= RS1("M_Qty_Safe")
		M_Qty_Include_coming	= RS1("M_Qty_Include_coming")
		M_Package_Unit				= RS1("M_Package_Unit")
	end if
	RS1.Close

	rem DB 업데이트
	RS1.Open "tbMaterial",sys_DBConString,3,2,2
	with RS1
		.AddNew
		.Fields("M_P_No")				= M_P_No
		.Fields("M_P_No_Sub")		= M_P_No_Sub
		.Fields("Partner_P_Name")	= Partner_P_Name
		.Fields("M_OSP_YN")			= M_OSP_YN
		.Fields("M_Desc")				= M_Desc
		.Fields("M_Spec")				= M_Spec
		.Fields("M_Price_Temp_YN") = M_Price_Temp_YN
		.Fields("M_Price_Apply_Date")	= date()
		.Fields("M_Price")			= cdbl(M_Price)
		.Fields("M_Package_Unit")		= M_Package_Unit
		.Fields("M_Price_LGE")	= cdbl(M_Price_LGE)
		.Fields("M_Process")		= M_Process
		.Fields("M_Maker")			= M_Maker
		.Fields("M_Division")		= M_Division
		.Fields("M_Qty")				= M_Qty
		.Fields("M_Qty_Safe")		= M_Qty_Safe
		.Fields("M_Qty_Include_coming")	= M_Qty_Include_coming
		.Fields("M_Reg_Date")		= date()
		.Fields("M_Reg_ID")			= gM_ID
		.Update
		.Close
	end with
	
	SQL = "select count(*) from tbMaterial where M_P_No = '"&M_P_No&"'"
	RS1.Open SQL,sys_DBCon
	if RS1(0) > 1 then
		SQL = "insert into tbMaterial_Price_Log (Material_M_P_No, Partner_P_Name, MPL_Temp_YN, MPL_Price, MPL_Price_LGE, MPL_Reg_Date, MPL_Apply_Date, MPL_Reg_ID, MPL_Desc) values "
		SQL = SQL & "('"&M_P_No&"','"&Partner_P_Name&"','"&M_Price_Temp_YN&"',"&cdbl(M_Price)&","&cdbl(M_Price_LGE)&",'"&date()&"','"&date()&"','"&gM_ID&"','복수거래처 등록')"
		sys_DBCon.execute(SQL)
	else
		SQL = "insert into tbMaterial_Price_Log (Material_M_P_No, Partner_P_Name, MPL_Temp_YN, MPL_Price, MPL_Price_LGE, MPL_Reg_Date, MPL_Apply_Date, MPL_Reg_ID, MPL_Desc) values "
		SQL = SQL & "('"&M_P_No&"','"&Partner_P_Name&"','"&M_Price_Temp_YN&"',"&cdbl(M_Price)&","&cdbl(M_Price_LGE)&",'"&date()&"','"&date()&"','"&gM_ID&"','신규자재')"
		sys_DBCon.execute(SQL)
	end if
	RS1.Close
	
	
end if

rem 객체 해제
Set RS1	= nothing
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="m_list.asp" method=post>
<input type="hidden" name="s_callby" value="mo_frame">
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="m_list.asp" method=post>
<input type="hidden" name="s_callby" value="mo_frame">
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