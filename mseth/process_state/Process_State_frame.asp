<!-- #include Virtual = "/mseth/header/asp_header.asp" -->
<!-- include Virtual = "/mseth/header/session_check_header.asp" -->
<!-- #include Virtual = "/mseth/header/db_header.asp" -->
<!-- #include Virtual = "/mseth/header/html_header.asp" -->
<!-- #include Virtual = "/mseth/header/layout_full_header.asp" -->
<!-- #include Virtual = "/mseth/header/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line

s_Work_Date	= Request("s_Work_Date")
s_Line		= Request("s_Line")

if s_Work_Date = "" then
	s_Work_Date = date()
end if

if s_Line = "" then
	s_Line = "1"
end if
%>

<script language="javascript">
function all_print()
{
	ifrmPR_Print1.focus();
	ifrmPR_Print1.print();
}
</script>

<table width=100px cellpadding=0 cellspacing=0 border=0>
<form name="frmSearch" action="process_state_frame.asp" method="post">
<tr>
	<td><input type="text" name="s_Work_Date" value="<%=s_Work_Date%>" style="width:80px;height:19px" onclick="Calendar_D(this)"></td>
	<td>
		<select name="s_Line" onchange="frmSearch.submit();">
		<option value="1"<%if s_Line="1" then%> selected<%end if%>>1라인</option>
		<option value="2"<%if s_Line="2" then%> selected<%end if%>>2라인</option>
		<option value="3"<%if s_Line="3" then%> selected<%end if%>>3라인</option>
		<option value="4"<%if s_Line="4" then%> selected<%end if%>>4라인</option>
		<option value="5"<%if s_Line="5" then%> selected<%end if%>>5라인</option>
		</select>
	</td>
	<td width=1px><img src="/img/blank.gif" width=1px height=1px></td>
	<td><input type="submit" value="조회">&nbsp;&nbsp;<input type="button" value="라인별모니터" onclick="javascript:window.open('mtr_process_state_record.asp?s_Work_Date=<%=s_Work_Date%>&s_Line=<%=s_Line%>')"><!--&nbsp;&nbsp;<input type="button" value="전라인모니터" onclick="javascript:window.open('mtr_process_state_all_record.asp?s_Work_Date=<%=s_Work_Date%>')">--></td>
</tr>
</form>
</table>
<center>
<br>
<table width=730px cellpadding=0 cellspacing=0 border=1>
<tr>
	<td width=730px>
		<%=s_Line%>라인 생산 계획<br>
		<div id="idPlan" style="width:100%;height:1920px;">
			<iframe name="ifrmPlan" src="process_state_plan.asp?s_Work_Date=<%=s_Work_Date%>&s_Line=<%=s_Line%>" frameborder=0 width=100% height=100% style="border:1px solid darkred"></iframe>
		</div>
	</td>
</tr>
</table>
</center>

<!-- include Virtual = "/mseth/header/layout_tail.asp" -->
<!-- #include Virtual = "/mseth/header/html_tail.asp" -->
<!-- #include Virtual = "/mseth/header/db_tail.asp" -->
<!-- include Virtual = "/mseth/header/session_check_tail.asp" -->