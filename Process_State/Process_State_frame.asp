<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line

s_Work_Date	= Request("s_Work_Date")
s_Line		= Request("s_Line")

if s_Work_Date = "" then
	s_Work_Date = date()
end if

if s_Line = "" then
	s_Line = "PCBA1"
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
		<option value="PCBA1"<%if s_Line="PCBA1" then%> selected<%end if%>>PCBA1</option>
		<option value="PCBA2"<%if s_Line="PCBA2" then%> selected<%end if%>>PCBA2</option>
		<option value="PCBA3"<%if s_Line="PCBA3" then%> selected<%end if%>>PCBA3</option>
		<option value="PCBA4"<%if s_Line="PCBA4" then%> selected<%end if%>>PCBA4</option>
		<option value="PCBA5"<%if s_Line="PCBA5" then%> selected<%end if%>>PCBA5</option>
		<option value="PCBA6"<%if s_Line="PCBA6" then%> selected<%end if%>>PCBA6</option>
		<option value="PCBA7"<%if s_Line="PCBA7" then%> selected<%end if%>>PCBA7</option>
		<option value="CBOX1"<%if s_Line="CBOX1" then%> selected<%end if%>>CBOX1</option>
		<option value="CBOX2"<%if s_Line="CBOX2" then%> selected<%end if%>>CBOX2</option>
		<option value="CBOX3"<%if s_Line="CBOX3" then%> selected<%end if%>>CBOX3</option>
		<option value="CBOX4"<%if s_Line="CBOX4" then%> selected<%end if%>>CBOX4</option>
		<option value="CBOX5"<%if s_Line="CBOX5" then%> selected<%end if%>>CBOX5</option>
		<option value="CBOX6"<%if s_Line="CBOX6" then%> selected<%end if%>>CBOX6</option>
		<option value="CBOX7"<%if s_Line="CBOX7" then%> selected<%end if%>>CBOX7</option>
		</select>
	</td>
	<td width=1px><img src="/img/blank.gif" width=1px height=1px></td>
	<td><input type="submit" value="조회">&nbsp;&nbsp;
		<input type="button" value="라인별현황판" onclick="javascript:window.open('mtr_process_state_record.asp?s_Work_Date=<%=s_Work_Date%>&s_Line=<%=s_Line%>')">&nbsp;&nbsp;
		<input type="button" value="PCBA통합현황판" onclick="javascript:window.open('mtr_process_state_all_record.asp?s_Process=PCBA&s_Work_Date=<%=s_Work_Date%>')">&nbsp;&nbsp;
		<input type="button" value="CBOX통합현황판" onclick="javascript:window.open('mtr_process_state_all_record.asp?s_Process=CBOX&s_Work_Date=<%=s_Work_Date%>')">&nbsp;&nbsp;
		<!--<input type="button" value="개발중" onclick="javascript:window.open('new_mtr_process_state_all_record.asp?s_Process=PCBA&s_Work_Date=<%=s_Work_Date%>')">&nbsp;&nbsp;-->
		<input type="button" value="멀티현황판" onclick="javascript:window.open('mtr_process_state_all_record.asp?s_Multi_YN=Y&s_Process=PCBA&s_Work_Date=<%=s_Work_Date%>')">
		</td>
</tr>
</form>
</table>
<center>
<br>
<table width=730px cellpadding=0 cellspacing=0 border=1>
<tr>
	<td width=730px>
		<%=s_Line%>라인 생산 계획<br>
		<div id="idPlan" style="width:100%;height:1900px;">
			<iframe name="ifrmPlan" src="process_state_plan.asp?s_Work_Date=<%=s_Work_Date%>&s_Line=<%=s_Line%>" frameborder=0 width=100% height=100% style="border:1px solid darkred"></iframe>
		</div>
	</td>
</tr>
</table>
</center>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->