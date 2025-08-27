<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim SQL
dim RS1
dim RS2

dim CNT1
dim CNT2
dim CNT3

dim strGongjung
dim arrGongjung
dim strLines
dim arrLines
dim arrLine

dim seedDate

dim am_pm
dim Time_Real

strGongjung		= "SMD;MAN"
strLines		= "Y-1:Y-2:Y-3;P-1:P-2:P-3"	

arrGongjung 	= split(strGongjung,";")
arrLines	 	= split(strLines,";")

seedDate = request("seedDate")
if seedDate = "" then
	seedDate = date()
end if
%>

<table width=600px cellspacing=0 border=1 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<form name="frmCalendar" method="post" action="#">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
<input type="hidden" name="">
</form>
<tr>
	<td width=40px>공정</td>
	<td width=70px>시간</td>
<%
for CNT1=0 to 6
%>
	<td width=80px><%=dateadd("d",CNT1,seedDate)%></td>
<%
next
%>
</tr>
<%
for CNT1=0 to ubound(arrGongjung)
%>
<tr>
	<td width=40px><%=arrGongjung(CNT1)%></td>
	<td width=70px>
		<table width=70px cellspacing=0 border=1 bgcolor="#ffffff"  style="table-layout:fixed" style="border-collapse:collapse">
<%
	for CNT2=8 to 36
		Time_Real = CNT2
		am_pm = "am"
		if Time_Real > 12 then
			am_pm = "pm"
		elseif Time_Real > 24 then
			am_pm = "am"
			Time_Real = Time_Real - 24
		end if
		if len(Time_Real) = 1 then
			Time_Real = "0" & Time_Real
		end if
%>
		<tr>
			<td width=70px><%=Time_Real%>:00(<%=am_pm%>)</td>
		</tr>
		<tr>
			<td width=70px><%=Time_Real%>:30(<%=am_pm%>)</td>
		</tr>
<%
	next
%>
		</table>
	</td>
<%
	for CNT2=0 to 6
%>
	<td width=80px>
		<table width=80px cellspacing=0 border=1 bgcolor="#ffffff"  style="table-layout:fixed" style="border-collapse:collapse">
		<tr>
<%
		arrLine = split(arrLines(CNT1),":")
		for CNT3=0 to ubound(arrLine)
%>
			<td><%=arrLine(CNT3)%></td>
<%
		next
%>
		</tr>
		</table>
	</td>
<%
	next
%>
</tr>
<%
next
%>
</table>

<!-- include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->