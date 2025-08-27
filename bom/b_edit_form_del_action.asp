<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include Virtual = "/function/inc_share_function.asp" -->
<% 
dim SQL
dim B_Code
dim B_D_No

B_Code = request("B_Code")
B_D_No = request("B_D_No")

dim strTable

dim RS1

SQL = "delete tbBOM_Qty where BOM_B_Code = "&B_Code
sys_DBCon.execute(SQL)

SQL = "delete tbBOM_Qty_Archive where BOM_B_Code = "&B_Code
sys_DBCon.execute(SQL)

SQL = "delete tbBOM_Sub where BOM_B_Code = "&B_Code
sys_DBCon.execute(SQL)

SQL = "delete tbBOM where B_Code = "&B_Code
sys_DBCon.execute(SQL)
%>
<form name="frmRedirect" action="b_list.asp" method=post>
<input type="hidden" name="postback_yn" value="Y">
</form>
<script language="javascript">
frmRedirect.submit();
</script>


<!-- #include Virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->