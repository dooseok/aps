<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->

<%
dim SQL

SQL = "update tbLGE_Model set LM_Company='타사' where LM_Company='미분류'"
sys_DBCon.execute(SQL)
%>

<form name="frmRedirect" method="post" action="lm_list.asp">
</form>
<script language="javascript">
frmRedirect.submit();
</script>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->