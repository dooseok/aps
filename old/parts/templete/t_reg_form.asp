<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<script language="javascript">
function frmNew_Check()
{
	if(frmNew.MT_Name.value=="")
	{
		alert("���ø� ������ �Է����ּ���.");
		return false;
	}
	else
	{
		frmNew.submit();
		return true;
	}
}
</script>
<br><br><br><br><br>
<center>
���ø��� �̸��� �Է����ּ���.
<form name="frmNew" action="t_reg_action.asp" method="post">
<input type="text" name="MT_Name" value="">&nbsp;
<input type="button" value="���" onclick="frmNew_Check()">
<input type="hidden" name="s_Opener_Type" value="<%=Request("s_Opener_Type")%>">
<input type="hidden" name="s_Opener_Code" value="<%=Request("s_Opener_Code")%>">
</form>
</center>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->
