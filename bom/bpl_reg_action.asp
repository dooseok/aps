<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL
dim RS1

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev		= Request("URL_Prev")
URL_Next		= Request("URL_Next")
		
rem 에러메세지가 있을 경우 실행안됨
if strError = "" then

	SQL = "insert into tbBOM_Price_Log (BOM_Sub_BS_D_No,BPL_Market,BPL_Currency,BPL_Price,Partner_P_Name,BPL_Apply_Date,BPL_Reg_Date) values "
	SQL = SQL & "('"&trim(Request("BOM_Sub_BS_D_No"))&"',"
	SQL = SQL & "'"&trim(Request("BPL_Market"))&"',"
	SQL = SQL & "'"&trim(Request("BPL_Currency"))&"',"
	SQL = SQL & "'"&trim(Request("BPL_Price"))&"',"
	SQL = SQL & "'"&trim(Request("Partner_P_Name"))&"',"
	SQL = SQL & "'"&trim(Request("BPL_Apply_Date"))&"',"
	SQL = SQL & "'"&date()&"')"
	sys_DBCon.execute(SQL)
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
<form name="frmRedirect" action="bpl_list.asp" method=post>

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
<form name="frmRedirect" action="bpl_list.asp" method=post>

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