<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<%
dim Request_Fields
dim strRequestForm
dim strRequestQueryString
for each Request_Fields in Request.Form
	'if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	'end if
next
for each Request_Fields in Request.QueryString
	'if lcase(left(Request_Fields,2))="s_" then
		strRequestForm = strRequestForm & "<input type='hidden' name='"&Request_Fields&"' value='"&Request(Request_Fields)&"'>" &vbcrlf
	'end if
next

dim URL_Back
dim strPK
dim strPK_Value

URL_Back	= Request("URL_Back")
strPK		= Request("strPK")
strPK_Value	= Request("strPK_Value")
%>

<form name="frmRedirect" action="<%=URL_Back%>" method=post>
<input type="hidden" name="<%=strPK%>" value="<%=strPK_Value%>">
<%
'response.write strRequestForm
%>
</form>

<script language="javascript">
alert("�� �������� �����Ͻ� �� �����ϴ�.\n������ �ʿ��մϴ�.");
frmRedirect.submit();
</script>



<!-- #include Virtual = "/header/db_tail.asp" -->