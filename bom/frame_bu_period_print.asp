<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<script language="javascript">
function frmCheck()
{
	frmSearch.submit();
}
</script>

<center>
<table width=700px cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed">
<form name="frmSearch" method="post" action="bu_period_print.asp" target="ifrmReport">
<tr>
	<td width=100% align=center style="font-size:12px;">
		[접수일]
		<input type="text" name="s_Date_1" class="input" style="width:80px" readonly onclick="javascript:Calendar_D(this);">
		 -
		<input type="text" name="s_Date_2" class="input" style="width:80px" readonly onclick="javascript:Calendar_D(this);">
		<input type="button" value="조회" onclick="javascript:frmCheck()">
		<input type="button" value="인쇄" onclick="javascript:ifrmReport.UsePrint()">

	</td>
</tr>
</form>
</table>
</center>

<iframe name="ifrmReport" width=100% height=100%></iframe>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->