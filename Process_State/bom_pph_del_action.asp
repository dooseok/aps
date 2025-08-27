<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->

<%
rem 변수선언
dim SQL

dim BP_Code

dim temp
dim strError
dim URL_Prev
dim URL_Next

URL_Prev	= Request("URL_Prev")
URL_Next	= Request("URL_Next")

BP_Code	= trim(Request("BP_Code"))

SQL = "delete tbBOM_PPH where BP_Code = "&BP_Code
sys_DBCon.execute(SQL)
%>

<%
if strError = "" then
%>
<form name="frmRedirect" action="bom_pph_list.asp" method=post>

</form>
<script language="javascript">
frmRedirect.submit();
</script>
<%
else
%>
<form name="frmRedirect" action="bom_pph_list.asp" method=post>

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