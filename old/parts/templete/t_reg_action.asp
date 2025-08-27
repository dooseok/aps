<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim strError

dim s_Opener_Type
dim s_Opener_Code
dim MT_Name

s_Opener_Type	= Request("s_Opener_Type")
s_Opener_Code	= Request("s_Opener_Code")
MT_Name			= Request("MT_Name")

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(MT_Name) from tbMaterial_Templete where MT_Name = '"&MT_Name&"'"
RS1.Open SQL,sys_DBCon
if RS1(0) > 0 then
	strError = "* 동일한 이름의 템플릿이 있습니다."
end if
RS1.Close
set RS1 = nothing

if strError = "" then
	SQL = "insert into tbMaterial_Templete (MT_Name) values ('"&MT_Name&"')"
	sys_DBCon.execute(SQL)
	if s_Opener_Type = "Transaction" then
		SQL = "insert into tbMaterial_Templete_detail select Material_Templete_MT_Name='"&MT_Name&"', MPD_Count=MTD_Qty, Material_M_P_No from tbMaterial_Transaction_Detail where Material_Transaction_MT_Code = "&s_Opener_Code
		sys_DBCon.execute(SQL)
	elseif s_Opener_Type = "Order" then
		SQL = "insert into tbMaterial_Templete_detail select Material_Templete_MT_Name='"&MT_Name&"', MPD_Count=MOD_Qty,Material_M_P_No from tbMaterial_Order_Detail where Material_Order_MO_Code = "&s_Opener_Code
		sys_DBCon.execute(SQL)
	end if
end if

if strError = "" then
%>
<form name="frmRedirect" action="t_frame.asp" method=post>
<input type="hidden" name="s_Opener_Type" value="<%=s_Opener_Type%>">
<input type="hidden" name="s_Opener_Code" value="<%=s_Opener_Code%>">
<input type="hidden" name="MT_Name" value="<%=MT_Name%>">
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="t_reg_form.asp" method=post>
<input type="hidden" name="s_Opener_Type" value="<%=s_Opener_Type%>">
<input type="hidden" name="s_Opener_Code" value="<%=s_Opener_Code%>">
<input type="hidden" name="MT_Name" value="<%=MT_Name%>">
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