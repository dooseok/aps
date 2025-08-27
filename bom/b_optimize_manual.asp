<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<Table border=1>
<%
dim RS1
dim SQL

set RS1 = server.createObject("ADODB.RecordSet")

SQL = "select count(B_Code) from tbBOM where (B_Opt_YN <> 'Y' or B_Opt_YN is null) and B_Version_Current_YN='Y'"
RS1.Open SQL,sys_DBCon

response.write "사용중인 최적화 안된 품번이 "&RS1(0)&"개 입니다.<br>"

RS1.Close

SQL = "select count(B_Code) from tbBOM where B_Opt_YN <> 'Y' or B_Opt_YN is null "
RS1.Open SQL,sys_DBCon

response.write "모든 최적화 안된 품번이 "&RS1(0)&"개 입니다."

RS1.Close
set RS1 = nothing
%>

</table>
<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->