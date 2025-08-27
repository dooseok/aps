<!-- #include virtual = "/header/asp_header.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<script language="javascript">
function faPrint()
{
	factory.printing.header				= "";
	factory.printing.footer				= "";
	factory.printing.portrait			= false;
	factory.printing.leftMargin			= 0.5;
	factory.printing.rightMargin		= 0.5;
	factory.printing.topMargin			= 0.5;
	factory.printing.bottomMargin		= 0.5;
	factory.printing.print(true, window);
}
</script>

<%
call usePrinter()
%>

<%
dim SQL
dim RS1
dim cntPage
dim cntRow

dim s_CR_Date_1
dim s_CR_Date_2
dim s_CR_Part


s_CR_Date_1	= Request("s_CR_Date_1")
s_CR_Date_2	= Request("s_CR_Date_2")
s_CR_Part	= Request("s_CR_Part")

set RS1 = server.CreateObject("ADODB.RecordSet")

SQL = ""
SQL = SQL & "select * from tbCost_Report where "
if s_CR_Part <> "" then
	SQL = SQL & "	CR_Part = '"&s_CR_Part&"' and "
end if
SQL = SQL & "	CR_Date between '"&s_CR_Date_1&"' and '"&s_CR_Date_2&"' order by CR_Code desc"
RS1.Open SQL,sys_DBCon

cntPage = 1
cntRow = 1
%>

<%
do until RS1.Eof
	if cntRow = 1 then
		if cntPage = 1 then
%>		
<table width=960px>
<tr>
	<td colspan=2 align=center style="font-size=30px"><b>지출금관리대장</b></td>
</tr>
</table>

<table width=960px cellpadding=0 cellspacing=0 border=1 bordercolor="gray" bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr bgcolor=skyblue>
	<td><b>부서별 도표</td>
	<td><b>용도별 도표</td>
</tr>
<tr height=280px>
	<td width=480px align=center valign=center><img name="imgChart1" src="/cost_report/temp/loading.gif"></td>
	<td width=480px align=center valign=center><img name="imgChart2" src="/cost_report/temp/loading.gif"></td>
</tr>
</table>
<br>
<%
		end if
%>
<table width=960px cellpadding=0 cellspacing=0 border=1 bordercolor="gray" bgcolor="#ffffff" style="border-collapse:collapse">
<tr bgcolor=skyblue>
	<td colspan=6><b>지출내역</td>
	<td colspan=3><b>결재</td>
	<td rowspan=2><b>비고</td>
</tr>
<tr bgcolor=skyblue>
	<td width=70px><b>부서</td>
	<td width=70px><b>일자</td>
	<td width=90px><b>거래선</td>
	<td width=130px><b>품명</td>
	<td width=180px><b>용도</td>
	<td width=80px><b>금액</td>
	<td width=50px><b>팀장</td>
	<td width=50px><b>이사</td>
	<td width=50px><b>대표</td>
</tr>
<%
	end if
%>
<tr>
	<td><%=RS1("CR_Part")%></td>
	<td><%=RS1("CR_Date")%></td>
	<td><%=RS1("CR_Line")%></td>
	<td><%=RS1("CR_Title")%></td>
	<td><%=RS1("CR_Use1")%><%if RS1("CR_Use2") <> "" then%> - <%end if%><%=RS1("CR_Use2")%></td>
	<td><%=RS1("CR_Amount")%></td>
	<td>
	<%
	if instr(ifrm_cr_chart_3,"-"&Request.Cookies("ADMIN")("M_ID")&"-") > 0 then
	%>
		<input type="checkbox" name="CR_Sign_TeamJang_YN"<%if RS1("CR_Sign_TeamJang_YN")="Y" then%> checked<%end if%> onclick="Sign_Save('<%=RS1("CR_Code")%>','CR_Sign_TeamJang_YN',this)"></td>
	<%
	Else
		if RS1("CR_Sign_TeamJang_YN")="Y" then
			response.write("V")
		end if
	End if
	%>
	<td>
	<%
	if instr(ifrm_cr_chart_2,"-"&Request.Cookies("ADMIN")("M_ID")&"-") > 0 then		
	%>
		<input type="checkbox" name="CR_Sign_Isa_YN"<%if RS1("CR_Sign_Isa_YN")="Y" then%> checked<%end if%> onclick="Sign_Save('<%=RS1("CR_Code")%>','CR_Sign_Isa_YN',this)"></td>
	<%
	Else
		if RS1("CR_Sign_Isa_YN")="Y" then
			response.write("V")
		end if
	End if
	%>
	<td>
	<%
	if instr(ifrm_cr_chart_1,"-"&Request.Cookies("ADMIN")("M_ID")&"-") > 0 then	
	%>	
		<input type="checkbox" name="CR_Sign_Sajang_YN"<%if RS1("CR_Sign_Sajang_YN")="Y" then%> checked<%end if%> onclick="Sign_Save('<%=RS1("CR_Code")%>','CR_Sign_SaJang_YN',this)"></td>
	<%
	Else
		if RS1("CR_Sign_Sajang_YN")="Y" then
			response.write("V")
		end if
	End if
	%>
	<td><%=RS1("CR_Memo")%></td>
</tr>
<%
	RS1.MoveNext
	if cntRow = 9000 then
%>
</table>
<%
		cntRow = 0
		cntPage = cntPage + 1
	end if	
	cntRow = cntRow + 1
loop
RS1.Close
%>
</table>


<iframe name="ifrmSign_Save" src="about:blank" height=0px height=0px frameborder=0></iframe>

<script language="javascript">
function Sign_Save(CR_Code,byWho,objSign)
{
	//var objCheck = eval("document."+byWho);
	//alert(objCheck);
	
	var toState = "N";
	
	if(objSign.checked)
	{
		toState = "Y";
	}
	ifrmSign_Save.location.href = "ifrm_cr_sign_save.asp?CR_Code=" + CR_Code + "&byWho=" + byWho + "&toState=" + toState;
}
</script>

<%
dim strCR_Part
dim strCR_Use1
dim strCR_Amount

SQL = ""
SQL = SQL & "select CR_Part, sum(CR_Amount) from tbCost_Report where "
if s_CR_Part <> "" then
	SQL = SQL & "	CR_Part = '"&s_CR_Part&"' and "
end if
SQL = SQL & "	CR_Date between '"&s_CR_Date_1&"' and '"&s_CR_Date_2&"' group by CR_Part order by CR_Part asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strCR_Part = strCR_Part & RS1(0) & ","
	strCR_Amount = strCR_Amount &  RS1(1) & ","
	RS1.MoveNext
loop
RS1.Close
if right(strCR_Part,1) = "," then
	strCR_Part = left(strCR_Part,len(strCR_Part)-1)
end if
if right(strCR_Amount,1) = "," then
	strCR_Amount = left(strCR_Amount,len(strCR_Amount)-1)
end if
'response.write strCR_Part & "<br>" & strCR_Amount & "<br>"
%>
<iframe src="/default.aspx?strKey=<%=strCR_Part%>&strValue=<%=strCR_Amount%>" frameborder=0 width=0 height=0></iframe>
<%
strCR_Use1		= ""
strCR_Amount	= ""
SQL = ""
SQL = SQL & "select CR_Use1, sum(CR_Amount) from tbCost_Report where "
if s_CR_Part <> "" then
	SQL = SQL & "	CR_Part = '"&s_CR_Part&"' and "
end if
SQL = SQL & "	CR_Date between '"&s_CR_Date_1&"' and '"&s_CR_Date_2&"' group by CR_Use1 order by CR_Use1 asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strCR_Use1 = strCR_Use1 & RS1(0) & ","
	strCR_Amount = strCR_Amount &  RS1(1) & ","
	RS1.MoveNext
loop
RS1.Close
if right(strCR_Use1,1) = "," then
	strCR_Use1 = left(strCR_Use1,len(strCR_Use1)-1)
end if
if right(strCR_Amount,1) = "," then
	strCR_Amount = left(strCR_Amount,len(strCR_Amount)-1)
end if
'response.write strCR_Use1 & "<br>" & strCR_Amount
%>
<iframe src="/default2.aspx?strKey=<%=strCR_Use1%>&strValue=<%=strCR_Amount%>" frameborder=0 width=0 height=0></iframe>
<%
set RS1 = nothing
%>

<script language="javascript">
function fRun()
{
	imgChart1.src = "/cost_report/temp/chart.png";
	imgChart2.src = "/cost_report/temp/chart2.png";
}
setTimeout("fRun()",5000);
</script>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->