<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<script language="javascript">
function FocusDown()
{
	if(frm2.ip1.value.length==13)
		frm2.ip2.focus();
}
function FocusUp()
{
	if(frm2.ip2.value.length==13)
		frm2.ip1.focus();
}
</script>

<table width=700px cellpadding=0 cellspacing=0 border=1>
<form name="frm2" method="post" action="">
<tr>
	<td><input type="text" name="ip1" onkeyup="javacript:FocusDown()"></td>
</tr>
<tr>
	<td><input type="text" name="ip2" onkeyup="javacript:FocusUp()"></td>
</tr>
</form>
</table>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->