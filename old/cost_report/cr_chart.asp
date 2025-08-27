<!-- #include virtual = "/header/asp_header.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_CR_Date_1
dim s_CR_Date_2
dim s_CR_Part

dim CNT1
dim strBasicDataPart
dim arrBasicDataPart
dim arrBasicDataPart2

if s_CR_Date_1 = "" and s_CR_Date_2 = "" then
	s_CR_Date_1 = dateadd("d",-7,date())
	s_CR_Date_2 = date()
elseif s_CR_Date_1 <> "" and s_CR_Date_2 = "" then
	s_CR_Date_2 = s_CR_Date_1
elseif s_CR_Date_1 = "" and s_CR_Date_2 <> "" then
	s_CR_Date_1 = s_CR_Date_2
elseif s_CR_Date_1 <> "" and s_CR_Date_2 <> "" then
end if
%>

<script language="javascript">
function all_print()
{
	ifrmCR_Chart1.focus();
	ifrmCR_Chart1.faPrint();
}
</script>

<table width=350px cellpadding=0 cellspacing=0 border=0>
<form name="frmSearch" action="ifrm_cr_chart.asp" method="post" target="ifrmCR_Chart1">
<tr>
	<td><input type="text" name="s_CR_Date_1" value="<%=s_CR_Date_1%>" style="width:65px;height:19px" onclick="Calendar_D(this)">부터</td>
	<td><input type="text" name="s_CR_Date_2" value="<%=s_CR_Date_2%>" style="width:65px;height:19px" onclick="Calendar_D(this)">까지</td>
	<td>
		<select name="s_CR_Part">
		<option value=""<%if s_CR_Part="" then%> selected<%end if%>>전체</option>
<%
strBasicDataPart = replace(BasicDataPart,"slt>","")
arrBasicDataPart = split(strBasicDataPart,";")

for CNT1 = 0 to ubound(arrBasicDataPart)
	arrBasicDataPart2 = split(arrBasicDataPart(CNT1),":")
%>
		<option value="<%=arrBasicDataPart2(0)%>"<%if s_CR_Part=arrBasicDataPart2(0) then%> selected<%end if%>><%=arrBasicDataPart2(1)%></option>
<%
next
%>
		</select>
	</td>
	<td width=1px><img src="/img/blank.gif" width=1px height=1px></td>
	<td><input type="submit" value="조회"><input type="button" value="인쇄" onclick="javascript:all_print();"></td>
</tr>
</form>
</table>
<br>
<iframe name="ifrmCR_Chart1" src="ifrm_cr_chart.asp?s_CR_Date_1=<%=s_CR_Date_1%>&s_CR_Date_2=<%=s_CR_Date_2%>&s_CR_Part=<%=s_CR_Part%>" frameborder=0 width=1000px height=1000px></iframe>


<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->