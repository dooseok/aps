<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
dim SQL

dim MTD_Code

dim strError

MTD_Code = request("MTD_Code")

dim MTD_Qty
dim Material_M_P_No

dim M_Qty
dim MT_Company

dim RS1
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select MTD_Qty, Material_M_P_No from tbMaterial_Transaction_Detail where MTD_Code='"&MTD_Code&"'"
RS1.Open SQL,sys_DBCon
MTD_Qty = 0
Material_M_P_No = ""
if not(RS1.Eof or RS1.Bof) then
	Material_M_P_No = RS1("Material_M_P_No")
	if isnumeric(RS1("MTD_Qty")) then
		MTD_Qty = RS1("MTD_Qty")
	end if
end if
RS1.Close


if Material_M_P_No <> "" then
	if Request("s_IpgoOrChulgo") = "Ipgo" then
		SQL = "update tbMaterial set M_Qty=M_Qty - "&MTD_Qty&" where M_P_No='"&Material_M_P_No&"'"
	elseif Request("s_IpgoOrChulgo") = "Chulgo" then
		SQL = "update tbMaterial set M_Qty=M_Qty + "&MTD_Qty&" where M_P_No='"&Material_M_P_No&"'"
	end if
	sys_DBCon.execute(SQL)
	
	SQL = "select top 1 MT_Company from tbMaterial_Transaction where MT_Code = "&Request("s_Material_Transaction_MT_Code")
	RS1.Open SQL,sys_DBCon
	MT_Company = RS1("MT_Company")
	RS1.Close
	
	SQL = "select top 1 M_Qty from tbMaterial where M_P_No='"&Material_M_P_No&"'"
	RS1.Open SQL,sys_DBCon
	M_Qty = RS1("M_Qty")
	RS1.Close
	
	SQL = "insert into tbMaterial_Stock_History (Material_M_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
	SQL = SQL &Material_M_P_No&"',"
	if Request("s_IpgoOrChulgo") = "Ipgo" then
		SQL = SQL &int(MTD_Qty) * -1&","
		SQL = SQL &M_Qty&",'"
		SQL = SQL &"사내입고취소','"
	elseif Request("s_IpgoOrChulgo") = "Chulgo" then
		SQL = SQL &MTD_Qty&","
		SQL = SQL &M_Qty&",'"
		SQL = SQL &"사내출고취소','"
	end if
	SQL = SQL &date()&"','"
	SQL = SQL &MT_Company&"')"
	
	if MTD_Qty = 0 then
	else
		sys_DBCon.execute(SQL)
	end if

end if

set RS1 = nothing

SQL = "delete from tbMaterial_Transaction_Detail where MTD_Code='"&MTD_Code&"'"
sys_DBCon.execute(SQL)
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
<form name="frmRedirect" action="MTD_list.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="MTD_list.asp" method=post>
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