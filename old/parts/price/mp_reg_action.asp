<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem ��������
dim SQL
dim RS1

dim MP_Price
dim Partner_P_Name

dim temp
dim strError
dim URL_Prev
dim URL_Next

MP_Price		= trim(Request("MP_Price"))
Partner_P_Name	= trim(Request("Partner_P_Name"))

URL_Prev = Request("URL_Prev")
URL_Next = Request("URL_Next")

rem �����޼����� ���� ��� ����ȵ�
if strError = "" then

	SQL = "insert into tbMaterial_Price (Material_M_P_No,MP_Price,Partner_P_Name) values "
	SQL = SQL & "('"&Request("s_Material_M_P_No")&"',"&MP_Price&",'"&Partner_P_Name&"')"

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
<form name="frmRedirect" action="MP_list.asp" method=post>
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
<form name="frmRedirect" action="MP_list.asp" method=post>
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