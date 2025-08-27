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
dim RS1
dim SQL

dim s_Date_1
dim s_Date_2

dim sumHeight

s_Date_1 = Request("s_Date_1")
s_Date_2 = Request("s_Date_2")

if trim(s_Date_1) = "" then
	s_Date_1 = dateadd("d",-7,date())
end if

if trim(s_Date_2) = "" then
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
	if(confirm("Ȯ���� Ŭ���Ͻ� �� ��ñ�ٸ��ø�\n�μ� ��ȭ���ڰ� ��ϴ�."))
	{
		factory.printing.print(true, window);
	}
}
</script>

<%
set RS1 = Server.CreateObject("ADODB.RecordSet")

SQL = ""
SQL = SQL & "select "&vbcrlf 
SQL = SQL & "	distinct MT_Company "&vbcrlf
SQL = SQL & "from tbMaterial_Transaction "&vbcrlf
SQL = SQL & "where "&vbcrlf
SQL = SQL & "	MT_Date between '"&s_Date_1&"' and '"&s_Date_2&"' and "&vbcrlf
SQL = SQL & "	exists "&vbcrlf
SQL = SQL & "		(select MTD_Code "&vbcrlf
SQL = SQL & "		from tbMaterial_Transaction_Detail "&vbcrlf
SQL = SQL & "		where "&vbcrlf
SQL = SQL & "			MTD_Qty > 0 and "&vbcrlf
if Request("strM_P_No") <> "" then
	SQL = SQL & "			Material_M_P_No = '"&Request("strM_P_No")&"' and "&vbcrlf
end if
SQL = SQL & "			Material_Transaction_MT_Code = MT_Code) "&vbcrlf
if Request("strMT_Company") <> "" then
	SQL = SQL & "	and MT_Company = '"&Request("strMT_Company")&"' "&vbcrlf
end if
SQL = SQL & "order by MT_Company asc "&vbcrlf

RS1.Open SQL,sys_DBCon
do until RS1.Eof
	sumHeight = Report_By_Company(RS1("MT_Company"),s_Date_1,s_Date_2)
%>
	<img src="/img/blank.gif" width=1px height="<%=1065 - (sumHeight mod 1065) + (int(sumHeight/1065) * 23)%>px"><br>
<%	
	RS1.MoveNext
loop
RS1.Close
set RS1 = nothing
%>

<%
function Report_By_Company(strMT_Company,s_Date_1,s_Date_2)
	dim RS1
	dim SQL
	dim CNT1
	dim Page
	dim MTD_Qty
	dim MTD_Price
	dim Qty_Price
	dim MT_Date
	dim oldMT_Date
	
	dim sum_MTD_Qty
	dim sum_Qty_Price
%>
<img src="/img/blank.gif" width=1px height=5px><br>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<tr height=60px>
	<td width=100% align=right style="font-size:20px;">
		<table class="pi_print_2" width=200px cellpadding=0 cellspacing=0 border=1 bgcolor="#333333" style="table-layout:fixed;">
		<tr bgcolor=white>
			<td width=30px rowspan=2>��<br>��</td>
			<td>�� ��</td>
			<td>�� ��</td>
			<td>�� ��</td>
			<td>�� ǥ</td>
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
<tr height=33px>
	<td width=100% align=center style="font-size:30px;">
		��������(��) ��ü�� ������
	</td>
</tr>
<tr height=15px>
	<td align=right style="font-size:12px;">
		<table width=100% cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td>[��ü��] : <%=strMT_Company%></td>
			<td width=350px>[��&nbsp;&nbsp;��]	<%=s_Date_1%>	 -	<%=s_Date_2%></td>
			<td width=150px align=left>[�����] : <%=date()%></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<%
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL = ""
	SQL = SQL & "select "&vbcrlf
	SQL = SQL & "	MT_Date = substring(convert(char(8),t1.MT_Date,112),5,2)+'.'+substring(convert(char(8),t1.MT_Date,112),7,2), "&vbcrlf
	SQL = SQL & "	t2.Material_M_P_No, "&vbcrlf
	SQL = SQL & "	M_Desc = (select M_Desc from tbMaterial where M_P_No = t2.Material_M_P_No), "&vbcrlf
	SQL = SQL & "	M_Spec = (select M_Spec from tbMaterial where M_P_No = t2.Material_M_P_No), "&vbcrlf
	SQL = SQL & "	M_Spec = (select M_Spec from tbMaterial where M_P_No = t2.Material_M_P_No), "&vbcrlf
	SQL = SQL & "	MTD_Price = convert(varchar(50),convert(decimal(10,2), "&vbcrlf
	SQL = SQL & "		isnull( "&vbcrlf
	SQL = SQL & "			(select top 1 MOD_Price from tbMaterial_Order_Detail where Material_M_P_No = t2.Material_M_P_No order by MOD_Code desc) "&vbcrlf
	SQL = SQL & "			, "&vbcrlf
	SQL = SQL & "			(select top 1 M_Price from tbMaterial where M_P_No = t2.Material_M_P_No)) "&vbcrlf
	SQL = SQL & "		)), "&vbcrlf
	SQL = SQL & "	Qty_Price = convert(varchar(50),convert(decimal(10,2),MTD_Qty * "&vbcrlf
	SQL = SQL & "		isnull( "&vbcrlf
	SQL = SQL & "			(select top 1 MOD_Price from tbMaterial_Order_Detail where Material_M_P_No = t2.Material_M_P_No order by MOD_Code desc) "&vbcrlf
	SQL = SQL & "			, "&vbcrlf
	SQL = SQL & "			(select top 1 M_Price from tbMaterial where M_P_No = t2.Material_M_P_No)) "&vbcrlf
	SQL = SQL & "		)), "&vbcrlf
	SQL = SQL & "	t2.MTD_Qty, "&vbcrlf
	SQL = SQL & "	t1.MT_Code "&vbcrlf
	SQL = SQL & "from "&vbcrlf
	SQL = SQL & "	tbMaterial_Transaction t1, "&vbcrlf
	SQL = SQL & "	tbMaterial_Transaction_Detail t2 "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	t1.MT_Code = t2.Material_Transaction_MT_Code and "&vbcrlf
	SQL = SQL & "	t1.MT_Company = '"&strMT_Company&"' and "&vbcrlf
	SQL = SQL & "	t1.MT_Date between '"&s_Date_1&"' and '"&s_Date_2&"' and "&vbcrlf
if Request("strM_P_No") <> "" then
	SQL = SQL & "		t2.Material_M_P_No = '"&Request("strM_P_No")&"' and "&vbcrlf
end if
	SQL = SQL & "	t1.MT_State = '���' "&vbcrlf

	SQL = SQL & "order by "&vbcrlf
	SQL = SQL & "	MT_Date asc "&vbcrlf
	
	RS1.Open SQL,sys_DBCon
	Page = 1	'������ �⺻��, 1
	CNT1 = 113	'���������� �⺻��, 0
	sum_MTD_Qty		= 0
	sum_Qty_Price	= 0
	
	oldMT_Date = RS1("MT_Date")
	
	do until RS1.Eof	'���ڵ�ݺ� [
		
		if CNT1 = 113 then		'ù ���� ���
%>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<tr bgcolor=black height=1><td width=700px colspan=8><img src="/img/black.gif" width=700px height=1px></td></tr>
</table>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<tr height=23px>
	<td align=center width=40px>�����</td>
	<td align=center width=120px>ǰ&nbsp;&nbsp;&nbsp;��</td>
	<td align=center width=100px>ǰ&nbsp;&nbsp;&nbsp;��</td>
	<td align=center>��&nbsp;&nbsp;&nbsp;��</td>
	<td align=right width=40px>��&nbsp;��</td>
	<td align=right width=60px>��&nbsp;��</td>
	<td align=right width=90px>��&nbsp;��</td>
	<td align=right width=70px>VAT</td>
	<td align=center width=50px>��ȣ</td>
</tr>
</table>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<tr bgcolor=black height=1><td width=700px colspan=8><img src="/img/black.gif" width=700px height=1px></td></tr>
</table>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<%
			CNT1 = CNT1 + 25	'�������	
	
		end if
		
		MT_Date			= RS1("MT_Date")
		MTD_Qty			= RS1("MTD_Qty")
		MTD_Price		= RS1("MTD_Price")
		Qty_Price		= RS1("Qty_Price")

		if oldMT_Date <> MT_Date then	'������� �ٲ� ��, ���� �������� ������� �����ֱ�
			'�Ұ�
%>
<tr height=23px>
	<td align=center colspan=2>* ��&nbsp;&nbsp;&nbsp;&nbsp;�� *</td>
	<td align=center>&nbsp;</td>
	<td align=center>&nbsp;</td>
	<td align=right><%=sum_MTD_Qty%></td>
	<td align=right>&nbsp;</td>
	<td align=right><%=FormatNumber(sum_Qty_Price, 2)%></td>
	<td align=right><%=FormatNumber(sum_Qty_Price * 0.1, 2)%></td>
	<td align=center>&nbsp;</td>
</tr>
</table>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<tr bgcolor=black height=1><td width=700px colspan=8><img src="/img/black.gif" width=700px height=1px></td></tr>
</table>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<%
			CNT1 = CNT1 + 24	'�������
			'������� ����
			sum_MTD_Qty		= 0
			sum_Qty_Price	= 0
		end if
%>
<tr height=23px>
	<td align=center width=40px nowrap><%=MT_Date%></td>
	<td align=center width=120px nowrap><%=RS1("Material_M_P_No")%></td>
	<td align=center width=100px nowrap><%=nLeft(RS1("M_Desc"),13)%></td>
	<td align=center nowrap><%=nLeft(RS1("M_Spec"),17)%></td>
	<td align=right width=40px nowrap><%=MTD_Qty%>&nbsp;</td>
	<td align=right width=60px nowrap><%=MTD_Price%>&nbsp;</td>
	<td align=right width=90px nowrap><%=FormatNumber(Qty_Price,2)%></td>
	<td align=right width=70px nowrap><%=FormatNumber(Qty_Price * 0.1,2)%></td>
	<td align=center width=50px nowrap><%=RS1("MT_Code")%></td>
</tr>
<%	
		CNT1 = CNT1 + 23	'�������
		sum_MTD_Qty		= sum_MTD_Qty		+ MTD_Qty
		sum_Qty_Price	= sum_Qty_Price	+ Qty_Price

		oldMT_Date = MT_Date '������ MT_Date�� �˱����� ����
		
		RS1.MoveNext		'�������ڵ��б�
	loop							'���ڵ�ݺ� ]
%>
<tr height=23px>
	<td align=center colspan=2>* ��&nbsp;&nbsp;&nbsp;&nbsp;�� *</td>
	<td align=center>&nbsp;</td>
	<td align=center>&nbsp;</td>
	<td align=right><%=sum_MTD_Qty%>&nbsp;</td>
	<td align=right>&nbsp;</td>
	<td align=right><%=FormatNumber(sum_Qty_Price, 2)%></td>
	<td align=right><%=FormatNumber(sum_Qty_Price * 0.1, 2)%></td>
	<td align=center>&nbsp;</td>
</tr>
</table>
<%
	CNT1 = CNT1 + 23	'�������
%>
<table width=700px cellpadding=0 cellspacing=0 border=0 bgcolor="#ffffff" style="table-layout:fixed;">
<tr bgcolor=black height=1><td width=700px colspan=8><img src="/img/black.gif" width=700px height=1px></td></tr>
</table>
<%
	CNT1 = CNT1 + 1
	set RS1 = nothing
	Report_By_Company = CNT1
end function
%>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->