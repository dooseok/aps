<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
call usePrinter()
%>

<%
dim CNT1
dim Page
dim RS1
dim SQL

dim s_Date_1
dim s_Date_2

dim Partner_P_Name
dim Ratio
dim NET_Price
dim VAT_Price
dim Price
dim P_Payment_Type

s_Date_1 = Request("s_Date_1")
s_Date_2 = Request("s_Date_2")

if s_Date_1 = "" then
	s_Date_1 = dateadd("d",-7,date())
end if

if s_Date_2 = "" then
	s_Date_2 = date()
end if
%>
<script language="javascript">
function UsePrint()
{
	factory.printing.header				= "";
	factory.printing.footer				= "";
	factory.printing.portrait			= true;
	factory.printing.leftMargin			= 0.5;
	factory.printing.rightMargin		= 0.5;
	factory.printing.topMargin			= 1;
	factory.printing.bottomMargin		= 1;
	if(confirm("확인을 클릭하신 후 잠시기다리시면\n인쇄 대화상자가 뜹니다."))
	{
		factory.printing.print(true, window);
	}
}
</script>

<%
SQL = ""

SQL = SQL & "select " &vbcrlf
SQL = SQL & "	Partner_P_Name, " &vbcrlf
SQL = SQL & "	Ratio = " &vbcrlf
SQL = SQL & "	case (select sum(POD_Price * POD_Qty) from tbParts_Order_Detail where POD_In_Date between '"&s_Date_1&"' and '"&s_Date_2&"') " &vbcrlf
SQL = SQL & "	when 0 then '0' " &vbcrlf
SQL = SQL & "	else " &vbcrlf
SQL = SQL & "		convert(varchar(50),convert(decimal(10,1), " &vbcrlf
SQL = SQL & "			sum(POD_Price * POD_Qty) " &vbcrlf
SQL = SQL & "			* 100 / " &vbcrlf
SQL = SQL & "			(select sum(POD_Price * POD_Qty) " &vbcrlf
SQL = SQL & "			from tbParts_Order_Detail " &vbcrlf
SQL = SQL & "			where POD_In_Date between '"&s_Date_1&"' and '"&s_Date_2&"'))) end, " &vbcrlf
SQL = SQL & "	NET_Price = round(sum(POD_Price * POD_Qty),0), " &vbcrlf
SQL = SQL & "	VAT_Price = round(sum(POD_Price * POD_Qty) * 0.1,0), " &vbcrlf
SQL = SQL & "	Price = round(sum(POD_Price * POD_Qty) + sum(POD_Price * POD_Qty) * 0.1,0), " &vbcrlf
SQL = SQL & "	P_Payment_Type = isnull((select top 1 (P_Payment_Type) from tbPartner where P_Name = Partner_P_Name),'') " &vbcrlf
SQL = SQL & "from " &vbcrlf
SQL = SQL & "	tbParts_Order t1, " &vbcrlf
SQL = SQL & "	tbParts_Order_Detail t2 " &vbcrlf
SQL = SQL & "where " &vbcrlf
SQL = SQL & "	t1.PO_Code = t2.Parts_Order_PO_Code and " &vbcrlf
SQL = SQL & "	POD_In_Date between '"&s_Date_1&"' and '"&s_Date_2&"' and POD_In_Qty > 0 " &vbcrlf
SQL = SQL & "group by " &vbcrlf
SQL = SQL & "	Partner_P_Name " &vbcrlf
SQL = SQL & "order by " &vbcrlf
SQL = SQL & "	Partner_P_Name " &vbcrlf

set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL,sys_DBCon

Page = 1
CNT1 = 0
do until RS1.Eof
	Partner_P_Name	= RS1("Partner_P_Name")
	Ratio			= RS1("Ratio")
	NET_Price		= replace(customformatcurrency(RS1("NET_Price")),"원","")
	VAT_Price		= replace(customformatcurrency(RS1("VAT_Price")),"원","")
	Price			= replace(customformatcurrency(RS1("Price")),"원","")
	P_Payment_Type	= RS1("P_Payment_Type")

	if CNT1 = 0 then
		if Page = 1  then
%>
<img src="/img/blank.gif" width=1px height=5px><br>
<table width=700px cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr height=60px>
	<td width=100% align=right style="font-size:20px;">
		<table class="pi_print_2" width=200px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed" style="border-collapse:collapse">
		<tr bgcolor=white>
			<td width=30px rowspan=2>결<br>제</td>
			<td>담 당</td>
			<td>이 사</td>
			<td>전 무</td>
			<td>대 표</td>
		</tr>
		<tr bgcolor=white height=40px>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td width=100% align=center style="font-size:30px;">
		엠에스이(주) 업체별 매입장 (요약)
	</td>
</tr>
<tr>
	<td align=right style="font-size:12px;">
		<table width=700px cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td width=150px>&nbsp;</td>
			<td width=400px>&nbsp;</td>
			<td width=150px align=left>Page : <%=Page%></td>
		</tr>
		<tr>
			<td width=150px>&nbsp;</td>
			<td width=400px>[기  간]	<%=s_Date_1%>	 -	<%=s_Date_2%></td>
			<td width=150px align=left>Date : <%=date()%></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<%
		else
%>
<img src="/img/blank.gif" width=1px height=5px><br>
<table width=700px cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed" style="border:none;">
<tr height=60px>
	<td width=100% align=right style="font-size:20px;">
		&nbsp;
	</td>
</tr>
<tr>
	<td width=100% align=center style="font-size:30px;">
		&nbsp;
	</td>
</tr>
<tr>
	<td width=100% align=right style="font-size:12px;">
		Page : <%=Page%><img src="/img/blank.gif" width=57px height=1px>&nbsp;<br>
		<img src="/img/blank.gif" width=130px height=1px>
		Date : <%=date()%>
	</td>
</tr>
</table>
<br>
<%
		end if
%>

<table class="pi_print_1" width=700px cellpadding=1 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed" style="border-collapse:collapse">
<tr bgcolor=black height=1><td colspan=6><img src="/img/black.gif" width=100% height=1px></td></tr>
<tr bgcolor=white>
	<td align=center>거 래 처</td>
	<td align=right>비율</td>
	<td align=right>공급  가액</td>
	<td align=right>세  액</td>
	<td align=right>매입  금액</td>
	<td align=center>결제방법</td>
</tr>
<tr bgcolor=black height=1><td colspan=6><img src="/img/black.gif" width=100% height=1px></td></tr>
<%
	end if
%>
<tr>
	<td><%=Partner_P_Name%></td>
	<td align=right><%=Ratio%></td>
	<td align=right><%=NET_Price%></td>
	<td align=right><%=VAT_Price%></td>
	<td align=right><%=Price%></td>
	<td><%=P_Payment_Type%></td>
<tr>
<%
	RS1.MoveNext
	CNT1 = CNT1 + 1
	if CNT1 = 36 and not(RS1.Eof or RS1.Bof) then
%>
<tr bgcolor=black height=1><td colspan=6><img src="/img/black.gif" width=100% height=1px></td></tr>
<tr bgcolor=white>
	<td colspan=6 align=right>[자료계속]<br><img src="/img/blank.gif" width=1px height=110px></td>
</tr>
<%
		page = page + 1
		CNT1 = 0
	end if

loop
RS1.Close
set RS1 = nothing
%>
<tr bgcolor=black height=1><td colspan=6><img src="/img/black.gif" width=100% height=1px></td></tr>
<tr bgcolor=white>
	<td colspan=6 align=right>[자료종료]</td>
</tr>
</table>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->