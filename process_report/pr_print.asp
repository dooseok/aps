<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_PR_Work_Date
dim s_PR_Work_Date2
dim s_Print_By

s_PR_Work_Date	= Request("s_PR_Work_Date")
s_PR_Work_Date2	= Request("s_PR_Work_Date2")
s_Print_By		= Request("s_Print_By")

if s_PR_Work_Date = "" then
	s_PR_Work_Date = dateadd("D",-1,date())
end if
if len(s_PR_Work_Date2) <> 22 then
	s_PR_Work_Date2 = ""
end if

if s_Print_By = "" then
	if Request.Cookies("ADMIN")("M_Part") = "제조1" then
		s_Print_By = "제조1"
	else
		s_Print_By = "SMD"
	end if
end if
%>

<script language="javascript">
function all_print()
{
	ifrmPR_Print1.focus();
	ifrmPR_Print1.print();
}
</script>

<table width=600px cellpadding=0 cellspacing=0 border=1>
<form name="frmSearch" action="pr_print.asp" method="post">
<tr>
	<td width=200px>기간:<input type="text" name="s_PR_Work_Date2" value="<%=left(s_PR_Work_Date2,10)%>" style="width:65px;height:19px" onclick="Calendar_D(this)">~<input type="text" name="s_PR_Work_Date2" value="<%=right(s_PR_Work_Date2,10)%>" style="width:65px;height:19px" onclick="Calendar_D(this)"></td>
	<td width=150px>특정일:<input type="text" name="s_PR_Work_Date" value="<%=s_PR_Work_Date%>" style="width:65px;height:19px" onclick="Calendar_D(this)"></td>
	<td>
		<select name="s_Print_By">
		<option value="제조1"<%if s_Print_By="제조1" then%> selected<%end if%>>제조1</option>
		<option value="SMD"<%if s_Print_By="SMD" then%> selected<%end if%>>SMD</option>
		<option value="IMD"<%if s_Print_By="IMD" then%> selected<%end if%>>IMD</option>
		<option value="DLV"<%if s_Print_By="DLV" then%> selected<%end if%>>영업</option>
		</select>
	</td>
	<td width=1px><img src="/img/blank.gif" width=1px height=1px></td>
	<td><input type="submit" value="조회"><input type="button" value="인쇄" onclick="javascript:all_print();"></td>
</tr>
</form>
</table>
<iframe name="ifrmPR_Print1" src="ifrm_pr_print.asp?s_PR_Work_Date=<%=s_PR_Work_Date%>&s_PR_Work_Date2=<%=s_PR_Work_Date2%>&s_Print_By=<%=s_Print_By%>" frameborder=1 width=1000px height=650px></iframe>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->