<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
			


<table width=280px cellpadding="1" cellspacing="1" align="center" border=0 bgcolor=#cccccc>
<tr bgcolor=white>
	<td colspan="6">
		<table width=280px cellpadding="0" cellspacing="0" align="center" border=0 bordercolor=gray>
		<tr>
			<td width=77px><%=Make_BTN("이전달","MovePreMonth()","")%></td>
			<td><b>실 &nbsp 적 &nbsp&nbsp&nbsp 현 &nbsp 황</b></td>
			<td width=77px><%=Make_BTN("다음달","MoveNextMonth()","")%></td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="yellow"> 
	<td width=80px><font color=blue><b>날짜＼부서</td>
	<td width=40px><font color=blue><b>IMD</td>
	<td width=40px><font color=blue><b>SMD</td>
	<td width=40px><font color=blue><b>수삽</td>
	<td width=40px><font color=blue><b>조립</td>
	<td width=40px><font color=blue><b>영업</td>
</tr>
			
			
<%
dim SQL
dim RS1

dim Today
dim MinDate
dim MaxDate

dim strMaxDate
dim arrMaxDate

Today = request("asrk")

if Today = "" then
	Today = date()
end if

strMaxDate = "31-28-31-30-31-30-31-31-30-31-30-31"
arrMaxDate = split(strMaxDate,"-")

MinDate = left(Today,8) & "01"
MaxDate = arrMaxDate(int(mid(Today,6,2))-1)
MaxDate = left(Today,8) & MaxDate

set RS1 = Server.CreateObject("adodb.recordset")


SQL = SQL & "select " &vbcrlf
SQL = SQL & "	distinct PR_Work_Date, " &vbcrlf
SQL = SQL & "	PR_Process, " &vbcrlf
SQL = SQL & "	sumPR_Amount = sum(PR_Amount) " &vbcrlf
SQL = SQL & "from tbProcess_Record " &vbcrlf
SQL = SQL & "	where" &vbcrlf
SQL = SQL & "		PR_Work_Date between  '" & MinDate & "'" &vbcrlf
SQL = SQL & "		and '" & MaxDate & "'" &vbcrlf
SQL = SQL & "group by " &vbcrlf
SQL = SQL & "	PR_Work_Date, " &vbcrlf
SQL = SQL & "	PR_Process " &vbcrlf
SQL = SQL & "order by PR_Work_Date asc" &vbcrlf
RS1.Open SQL,sys_DBCon

dim arrPR_Amount(31,4)
dim nDate

for nDate = 0 to ubound(arrPR_Amount)
	arrPR_Amount(nDate,0) = 0
	arrPR_Amount(nDate,1) = 0
	arrPR_Amount(nDate,2) = 0
	arrPR_Amount(nDate,3) = 0
	arrPR_Amount(nDate,4) = 0
next

do until RS1.Eof
	nDate = int(right(RS1("PR_Work_Date"),2))
	
	select case RS1("PR_Process")
		case "IMD"
			arrPR_Amount(nDate,0) = RS1("sumPR_Amount")
		case "SMD"
			arrPR_Amount(nDate,1) = RS1("sumPR_Amount")
		case "MAN"
			arrPR_Amount(nDate,2) = RS1("sumPR_Amount")
		case "ASM"
			arrPR_Amount(nDate,3) = RS1("sumPR_Amount")
		case "DLV"
			arrPR_Amount(nDate,4) = RS1("sumPR_Amount")
	end select
	
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing

for nDate = 1 to int(right(MaxDate,2))
%>
<tr <%if Weekday(left(Today,8) & nDate)=7 then%> bgcolor="skyblue"<%end if%><%if Weekday(left(Today,8) & nDate)=1 then%> bgcolor="pink"<%else%> bgcolor="white"<%end if%>>
	<td><%=left(today,8)%><%if len(nDate)=1 then%>0<%end if%><%=nDate%></td>
	<td><%=arrPR_Amount(nDate,0)%></td>
	<td><%=arrPR_Amount(nDate,1)%></td>
	<td><%=arrPR_Amount(nDate,2)%></td>
	<td><%=arrPR_Amount(nDate,3)%></td>
	<td><%=arrPR_Amount(nDate,4)%></td>
</tr>
<%
next
%>
</table>
	
<script language="javascript">
function MoveNextMonth()
{
	location.href="check.asp?asrk=<%=dateadd("m",+1,Today)%>";
}
function MovePreMonth()
{
	location.href="check.asp?asrk=<%=dateadd("m",-1,Today)%>";
}
</script>





<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->











