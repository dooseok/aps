<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->

<%
dim SQL
dim RS1

dim ER_Code

dim strError

dim ER_File_1
dim ER_File_2
dim ER_File_3

ER_Code = request("ER_Code")


set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbError_Reporting where ER_Code='"&ER_Code&"'"
RS1.Open SQL,sys_DBCon
ER_File_1		= RS1("ER_File_1")
ER_File_2		= RS1("ER_File_2")
ER_File_3		= RS1("ER_File_3")
File_Delete(DefaultPath_Error_Reporting & ER_File_1)
File_Delete(DefaultPath_Error_Reporting & ER_File_2)
File_Delete(DefaultPath_Error_Reporting & ER_File_3)
RS1.Close
set RS1 = nothing

SQL = "delete from tbError_Reporting where ER_Code='"&ER_Code&"'"
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
<form name="frmRedirect" action="ER_list.asp" method=post>

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
<form name="frmRedirect" action="ER_list.asp" method=post>

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