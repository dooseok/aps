<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim SQL
dim CNT1

dim Company
dim strLP_Model

dim arrLP_Model
dim LP_Model

Company		= Request("Company")
strLP_Model	= Request("strLP_Model")

arrLP_Model = split(strLP_Model,", ")

for CNT1 = 0 to ubound(arrLP_Model)
	LP_Model = trim(arrLP_Model(CNT1))
	SQL = "update tbLGE_Model set LM_Company='"&Company&"' where LM_Name='"&LP_Model&"'"
	sys_DBCon.execute(SQL)
next
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

%>
<form name="frmRedirect" action="lp_view.asp" method=post>
<%
response.write strRequestForm
%>
</form>
<script language="javascript">
frmRedirect.submit();
</script>


<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->