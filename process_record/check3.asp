<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
			


			<table width=280px cellpadding="0" cellspacing="0" align="center" border=1 bordercolor=gray>
			<tr>
				<td colspan="6">
					<table width=280px cellpadding="0" cellspacing="0" align="center" border=0 bordercolor=gray>
					<tr>
						<td width=77px><%=Make_BTN("������","MovePreMonth()","")%></td>
						<td><b>�� &nbsp �� &nbsp&nbsp&nbsp �� &nbsp Ȳ</b></td>
						<td width=77px><%=Make_BTN("������","MoveNextMonth()","")%></td>
					</tr>
					</table>
				</td>
			</tr>
			<tr bgcolor="yellow"> 
				<td width=80px><font color=blue><b>��¥���μ�</td>
				<td width=40px><font color=blue><b>IMD</td>
				<td width=40px><font color=blue><b>SMD</td>
				<td width=40px><font color=blue><b>����</td>
				<td width=40px><font color=blue><b>����</td>
				<td width=40px><font color=blue><b>����</td>
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


SQL = "select" &vbcrlf
SQL = SQL & "		distinct" &vbcrlf
SQL = SQL & "		PR_Work_Date ," &vbcrlf
SQL = SQL & "		IMD_���� = isnull ((select sum(PR_amount) from tbProcess_Record t2 where t2.PR_Work_Date = t1.PR_Work_Date and PR_Process='IMD'),'0')," &vbcrlf
SQL = SQL & "		SMD_���� = isnull ((select sum(PR_amount) from tbProcess_Record t2 where t2.PR_Work_Date = t1.PR_Work_Date and PR_Process='SMD'),'0')," &vbcrlf
SQL = SQL & "		MAN_���� = isnull ((select sum(PR_amount) from tbProcess_Record t2 where t2.PR_Work_Date = t1.PR_Work_Date and PR_Process='MAN'),'0')," &vbcrlf
SQL = SQL & "		DLV_���� = isnull ((select sum(PR_amount) from tbProcess_Record t2 where t2.PR_Work_Date = t1.PR_Work_Date and PR_Process='DLV'),'0')," &vbcrlf
SQL = SQL & "		ASM_���� = isnull ((select sum(PR_amount) from tbProcess_Record t2 where t2.PR_Work_Date = t1.PR_Work_Date and PR_Process='ASM'),'0')" &vbcrlf
SQL = SQL & "	from" &vbcrlf
SQL = SQL & "		tbProcess_Record t1" &vbcrlf
SQL = SQL & "	where" &vbcrlf
SQL = SQL & "		PR_Work_Date between  '" & MinDate & "'" &vbcrlf
SQL = SQL & "		and '" & MaxDate & "'" &vbcrlf
SQL = SQL & " order by PR_Work_Date asc" &vbcrlf
RS1.Open SQL,sys_DBCon
'response.write SQL

dim oldDate '������¥ �ӽú�����
dim skipDate
dim CNT

oldDate = MinDate
do until RS1.eof
	
	for CNT = 2 to datediff("d",oldDate,RS1("PR_Work_Date"))
			skipDate = dateadd("d",CNT-1,oldDate)
%>
			<tr <%if Weekday(skipDate)=7 then%> bgcolor="skyblue"<%end if%><%if Weekday(skipDate)=1 then%> bgcolor="pink"<%end if%>>
				<td><%=skipDate%></td>
				<td>0</td><td>0</td><td>0</td><td>0</td><td>0</td>
			</tr>
<%
	next
%>

			<tr <%if Weekday(RS1("PR_Work_Date"))=7 then%> bgcolor="skyblue"<%end if%><%if Weekday(RS1("PR_Work_Date"))=1 then%> bgcolor="pink"<%end if%>>
				<td><%= RS1("PR_Work_Date")%></td>
				<td><%= RS1("IMD_����")%></td>
				<td><%= RS1("SMD_����")%></td>
				<td><%= RS1("MAN_����")%></td>
				<td><%= RS1("DLV_����")%></td>
				<td><%= RS1("ASM_����")%></td>
			</tr>

<%
	'response.write RS1("PR_Work_Date") & ", " & RS1("IMD_����") & ", " & RS1("SMD_����") & ", " & RS1("MAN_����") & ", " & RS1("DLV_����") & ", " & RS1("ASM_����") & "<br>"

	oldDate = RS1("PR_Work_Date")
	RS1.MoveNext
loop


RS1.Close
set RS1 = nothing 




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











