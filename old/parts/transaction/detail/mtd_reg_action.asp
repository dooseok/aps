<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim Material_M_P_No
dim MTD_Qty
dim MTD_Remark
dim MTD_Ipgo_Date

dim temp
dim strError
dim URL_Prev
dim URL_Next

Material_M_P_No	= trim(Request("Material_M_P_No"))
MTD_Qty			= trim(Request("MTD_Qty"))
MTD_Remark		= trim(Request("MTD_Remark"))
MTD_Ipgo_Date	= trim(Request("MTD_Ipgo_Date"))

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

dim MT_Company
dim M_Qty

rem 에러메세지가 있을 경우 실행안됨
if strError = "" then
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	
	if Request("s_IpgoOrChulgo") = "Ipgo" then
		SQL = "update tbMaterial set M_Qty=M_Qty + "&MTD_Qty&" where M_P_No='"&Material_M_P_No&"'"
	elseif Request("s_IpgoOrChulgo") = "Chulgo" then
		SQL = "update tbMaterial set M_Qty=M_Qty - "&MTD_Qty&" where M_P_No='"&Material_M_P_No&"'"
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
		SQL = SQL &MTD_Qty&","
		SQL = SQL &M_Qty&",'"
		SQL = SQL &"사내입고','"
	elseif Request("s_IpgoOrChulgo") = "Chulgo" then
		SQL = SQL &int(MTD_Qty) * -1 &","
		SQL = SQL &M_Qty&",'"
		SQL = SQL &"사내출고','"
	end if
	SQL = SQL &date()&"','"
	SQL = SQL &MT_Company&"')"
	
	if MTD_Qty = 0 then
	else
		sys_DBCon.execute(SQL)
	end if

	if MTD_Ipgo_Date <> "" then
		SQL = "insert into tbMaterial_Transaction_Detail (Material_Transaction_MT_Code,Material_M_P_No,MTD_Qty,MTD_Remark,MTD_Ipgo_Date) values "
		SQL = SQL & "("&Request("s_Material_Transaction_MT_Code")&",'"&Material_M_P_No&"',"&MTD_Qty&",'"&MTD_Remark&"','"&MTD_Ipgo_Date&"')"
	else
		SQL = "insert into tbMaterial_Transaction_Detail (Material_Transaction_MT_Code,Material_M_P_No,MTD_Qty,MTD_Remark) values "
		SQL = SQL & "("&Request("s_Material_Transaction_MT_Code")&",'"&Material_M_P_No&"',"&MTD_Qty&",'"&MTD_Remark&"')"
	end if
	sys_DBCon.execute(SQL)
	
	set RS1 = nothing
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