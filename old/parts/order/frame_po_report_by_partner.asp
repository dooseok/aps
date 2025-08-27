<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
call Parts_Guide()
%>

<script language="javascript">
function frmCheck()
{
	frmSearch.submit();
}
</script>

<center>
<table width=700px cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed">
<form name="frmSearch" method="post" action="po_report_by_partner.asp" target="ifrmReport">
<tr>
	<td width=100% align=center style="font-size:12px;">
		[거래처]
		<select name="strPartner_P_Name">
			<option value="">전체</option>
<%
dim RS1
dim SQL

set RS1 = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from tbPartner order by P_Name asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
%>
	<option value="<%=RS1("P_Name")%>"><%=RS1("P_Name")%></option>
<%
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing
%>
		</select>
		&nbsp;
		[부  품]
		<input type="text" name="strP_P_No" class="input" style="width:120px" onclick="javascript:show_Parts_Guide(this);">
		&nbsp;
		[기  간]
		<input type="text" name="s_Date_1" class="input" style="width:80px" onclick="javascript:Calendar_D(this);">
		 -
		<input type="text" name="s_Date_2" class="input" style="width:80px" onclick="javascript:Calendar_D(this);">
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