<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
dim SQL

dim POD_Code

dim strError

POD_Code = request("POD_Code")

dim POD_In_Qty
dim Parts_P_P_No

dim M_Qty
dim Parts_Order_PO_Code
dim Partner_P_Name

dim RS1
set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select POD_In_Qty, Parts_P_P_No, Parts_Order_PO_Code from tbParts_Order_Detail where POD_Code='"&POD_Code&"'"
RS1.Open SQL,sys_DBCon
POD_In_Qty = 0
Parts_P_P_No = ""
if not(RS1.Eof or RS1.Bof) then
	Parts_P_P_No = RS1("Parts_P_P_No")
	Parts_Order_PO_Code = RS1("Parts_Order_PO_Code")
	if isnumeric(RS1("POD_In_Qty")) then
		POD_In_Qty = RS1("POD_In_Qty")
	end if
end if
RS1.Close

if Parts_P_P_No <> "" then
	SQL = 		"update tbParts set "
	SQL = SQL & "M_Qty=M_Qty - "&POD_In_Qty&" where P_P_No='"&Parts_P_P_No&"'"
	sys_DBCon.execute(SQL)
	
	SQL = "select M_Qty from tbParts where P_P_No='"&Parts_P_P_No&"'"
	RS1.Open SQL,sys_DBCon
	M_Qty = RS1("M_Qty")
	RS1.Close
	
	SQL = "select Partner_P_Name from tbParts_Order where PO_Code='"&Parts_Order_PO_Code&"'"
	RS1.Open SQL,sys_DBCon
	Partner_P_Name = RS1("Partner_P_Name")
	RS1.Close
	
	SQL = "insert into tbParts_Stock_History (Parts_P_P_No,MSH_Change_Stock,MSH_Applyed_Stock,MSH_Change_Type,MSH_Change_Date,MSH_Company) values ('"
	SQL = SQL &Parts_P_P_No&"',"
	SQL = SQL &POD_In_Qty * -1&","
	SQL = SQL &M_Qty&",'"
	SQL = SQL &"발주입고취소','"
	SQL = SQL &date()&"','"
	SQL = SQL &Partner_P_Name&"')"
	if POD_In_Qty <> 0 then
		sys_DBCon.execute(SQL)
	end if
end if

set RS1 = nothing

SQL = "delete from tbParts_Order_Detail where POD_Code='"&POD_Code&"'"
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
<form name="frmRedirect" action="POD_list.asp" method=post>
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
<form name="frmRedirect" action="POD_list.asp" method=post>
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